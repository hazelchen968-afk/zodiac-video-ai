#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
scene_splitter.py
用于将生肖运势长文案拆分为可执行的分镜表（JSON + Excel）。
"""

from __future__ import annotations

import json
import math
import re
import sys
from collections import Counter
from dataclasses import dataclass, asdict
from datetime import datetime
from pathlib import Path
from typing import Dict, List, Tuple

try:
    from openpyxl import Workbook
except ImportError:  # 仅在未安装 openpyxl 时提示用户
    Workbook = None


# 生肖英文标识到中文名映射
ANIMAL_NAME_ZH = {
    "rat": "鼠",
    "ox": "牛",
    "tiger": "虎",
    "rabbit": "兔",
    "dragon": "龙",
    "snake": "蛇",
    "horse": "马",
    "goat": "羊",
    "monkey": "猴",
    "rooster": "鸡",
    "dog": "狗",
    "pig": "猪",
}

# 生肖关键词（用于识别句子中提到的生肖）
ANIMAL_KEYWORDS = {
    "rat": ["鼠", "属鼠", "子鼠", "rat"],
    "ox": ["牛", "属牛", "丑牛", "ox"],
    "tiger": ["虎", "属虎", "寅虎", "tiger"],
    "rabbit": ["兔", "属兔", "卯兔", "rabbit"],
    "dragon": ["龙", "属龙", "辰龙", "dragon"],
    "snake": ["蛇", "属蛇", "巳蛇", "snake"],
    "horse": ["马", "属马", "午马", "horse"],
    "goat": ["羊", "属羊", "未羊", "goat", "ram"],
    "monkey": ["猴", "属猴", "申猴", "monkey"],
    "rooster": ["鸡", "属鸡", "酉鸡", "rooster"],
    "dog": ["狗", "属狗", "戌狗", "dog"],
    "pig": ["猪", "属猪", "亥猪", "pig"],
}

# 画面类型判定关键词
STYLE_RULES = [
    ("运势图表", ["财运", "收入", "投资"], "infographic, gold coins, data visualization"),
    ("实景氛围", ["事业", "工作", "升职"], "office scene, professional setting"),
    ("生肖形象", ["感情", "桃花", "婚姻"], "romantic scene, flowers, warm lighting"),
    ("实景氛围", ["健康", "身体"], "wellness, nature, fresh air"),
]

# 动态方式轮换，避免全部镜头重复
MOTION_CYCLE = ["轻微缩放", "横向平移", "上下浮动", "固定", "渐隐渐现"]


@dataclass
class Scene:
    """单条分镜结构。"""

    scene_id: str
    narration: str
    visual_prompt: str
    image_style: str
    zodiac_animal: str
    duration: int
    motion: str


class SceneSplitterError(Exception):
    """统一业务异常，便于输出友好提示。"""


def read_text_file(path: Path) -> str:
    """读取 UTF-8 文本并返回内容，处理不存在与编码错误。"""
    if not path.exists():
        raise SceneSplitterError(f"未找到输入文案文件：{path}")
    try:
        text = path.read_text(encoding="utf-8")
    except UnicodeDecodeError as exc:
        raise SceneSplitterError(f"文案文件编码错误，请确保为 UTF-8：{path}") from exc
    if not text.strip():
        raise SceneSplitterError("文案文件为空，请在 script.txt 中填写内容。")
    return text.strip()


def read_config(path: Path) -> Dict:
    """读取并校验配置 JSON。"""
    if not path.exists():
        raise SceneSplitterError(f"未找到配置文件：{path}")
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
    except UnicodeDecodeError as exc:
        raise SceneSplitterError(f"配置文件编码错误，请确保为 UTF-8：{path}") from exc
    except json.JSONDecodeError as exc:
        raise SceneSplitterError(f"config.json 格式错误：{exc}") from exc

    zodiac_animals = data.get("zodiac_animals")
    primary_animal = data.get("primary_animal")
    video_theme = data.get("video_theme", "生肖运势")

    if not isinstance(zodiac_animals, list) or not zodiac_animals:
        raise SceneSplitterError("config.json 中 zodiac_animals 必须是非空数组。")

    invalid = [a for a in zodiac_animals if a not in ANIMAL_NAME_ZH]
    if invalid:
        raise SceneSplitterError(f"zodiac_animals 存在不支持的生肖：{invalid}")

    if primary_animal not in ANIMAL_NAME_ZH:
        raise SceneSplitterError("primary_animal 不合法，请使用英文生肖标识。")

    if primary_animal not in zodiac_animals:
        raise SceneSplitterError("primary_animal 必须包含在 zodiac_animals 中。")

    return {
        "zodiac_animals": zodiac_animals,
        "primary_animal": primary_animal,
        "video_theme": str(video_theme),
    }


def split_sentences(text: str) -> List[str]:
    """优先按句号/问号/感叹号拆句。"""
    parts = re.split(r"(?<=[。！？!?])\s*", text)
    sentences = [p.strip() for p in parts if p and p.strip()]
    return sentences


def split_long_sentence(sentence: str, max_len: int = 120) -> List[str]:
    """仅当句子过长时，按逗号进一步切分。"""
    if len(sentence) <= max_len:
        return [sentence]

    chunks: List[str] = []
    current = ""
    for part in re.split(r"(?<=[，,])", sentence):
        if len(current) + len(part) <= max_len:
            current += part
        else:
            if current:
                chunks.append(current.strip())
            current = part
    if current.strip():
        chunks.append(current.strip())
    return chunks


def build_segments(text: str) -> List[str]:
    """构建初步片段，再合并到 50-120 字区间。"""
    base_units: List[str] = []
    for sentence in split_sentences(text):
        base_units.extend(split_long_sentence(sentence, 120))

    segments: List[str] = []
    buffer = ""
    for unit in base_units:
        if not buffer:
            buffer = unit
            continue

        if len(buffer) < 50:
            if len(buffer) + len(unit) <= 120:
                buffer += unit
            else:
                segments.append(buffer.strip())
                buffer = unit
        else:
            # 缓冲已达下限时，优先结束当前镜头保持语义完整
            segments.append(buffer.strip())
            buffer = unit

    if buffer.strip():
        if segments and len(buffer) < 50 and len(segments[-1]) + len(buffer) <= 120:
            segments[-1] += buffer
        else:
            segments.append(buffer.strip())

    return segments


def rebalance_scene_count(segments: List[str], target_min: int = 35, target_max: int = 55) -> List[str]:
    """通过拆分/合并将镜头数量调整到 35-55。"""
    adjusted = segments[:]

    # 数量不足：优先拆分较长片段
    while len(adjusted) < target_min:
        idx = max(range(len(adjusted)), key=lambda i: len(adjusted[i]))
        seg = adjusted[idx]
        if len(seg) <= 70:
            break
        split_pos = len(seg) // 2
        cut_candidates = [m.start() + 1 for m in re.finditer(r"[，,。；;]", seg)]
        if cut_candidates:
            split_pos = min(cut_candidates, key=lambda x: abs(x - split_pos))
        left, right = seg[:split_pos].strip(), seg[split_pos:].strip()
        if left and right:
            adjusted[idx:idx + 1] = [left, right]
        else:
            break

    # 数量过多：合并短片段
    while len(adjusted) > target_max:
        merged = False
        for i in range(len(adjusted) - 1):
            if len(adjusted[i]) + len(adjusted[i + 1]) <= 120:
                adjusted[i:i + 2] = [adjusted[i] + adjusted[i + 1]]
                merged = True
                break
        if not merged:
            break

    return adjusted


def detect_animal(text: str, configured_animals: List[str], default_animal: str) -> str:
    """检测片段中明确提及的生肖，否则返回主生肖。"""
    for animal in configured_animals:
        for kw in ANIMAL_KEYWORDS.get(animal, []):
            if kw.lower() in text.lower():
                return animal
    return default_animal


def detect_image_style(text: str) -> Tuple[str, str]:
    """根据关键词判断画面类型，并返回附加 prompt 元素。"""
    for style, keywords, prompt_extra in STYLE_RULES:
        if any(k in text for k in keywords):
            return style, prompt_extra
    return "文字卡片", "typography, calligraphy, clean background"


def calculate_duration(text: str, image_style: str) -> int:
    """按规则计算时长。"""
    if image_style == "运势图表":
        return 4

    char_len = len(text)
    duration = math.ceil(char_len / 18)
    if char_len < 30:
        duration = max(duration, 3)
    if char_len > 100:
        duration = min(duration, 8)
    return max(duration, 1)


def build_visual_prompt(animal: str, style_extra: str, theme: str) -> str:
    """生成 60-80 词英文 prompt。"""
    animal_en = animal
    prompt = (
        f"An anthropomorphic {animal_en} character as the main subject for {theme}, "
        f"expressive and cinematic composition, {style_extra}, consistent character design, "
        f"same anthropomorphic {animal_en} character across scenes. Chinese New Year illustration style, "
        f"highly detailed, festive ornaments and symbolic elements, red and gold color scheme with balanced "
        f"contrast, soft volumetric light, 16:9 aspect ratio, 4K, best quality."
    )
    words = len(prompt.split())
    if words < 60:
        prompt += " Intricate textures, layered background depth, polished rendering, dynamic yet clear storytelling focus."
    elif words > 80:
        # 简单裁剪到约 80 词
        prompt = " ".join(prompt.split()[:80])
    return prompt


def make_output_dir(base: Path) -> Path:
    """创建按日期归档的输出目录，已存在则自动追加序号。"""
    date_str = datetime.now().strftime("%Y%m%d")
    candidate = base / date_str
    if not candidate.exists():
        candidate.mkdir(parents=True, exist_ok=False)
        return candidate

    idx = 1
    while True:
        candidate = base / f"{date_str}_{idx:02d}"
        if not candidate.exists():
            candidate.mkdir(parents=True, exist_ok=False)
            return candidate
        idx += 1


def scenes_to_excel(scenes: List[Scene], output_path: Path) -> None:
    """写入 Excel 分镜表。"""
    if Workbook is None:
        raise SceneSplitterError(
            "未安装 openpyxl，无法输出 xlsx。请先执行：pip install openpyxl"
        )

    wb = Workbook()
    ws = wb.active
    ws.title = "storyboard"

    headers = [
        "scene_id",
        "narration",
        "visual_prompt",
        "image_style",
        "zodiac_animal",
        "duration",
        "motion",
    ]
    ws.append(headers)

    for scene in scenes:
        ws.append([
            scene.scene_id,
            scene.narration,
            scene.visual_prompt,
            scene.image_style,
            scene.zodiac_animal,
            scene.duration,
            scene.motion,
        ])

    wb.save(output_path)


def print_stats(scenes: List[Scene], configured_animals: List[str]) -> None:
    """打印统计信息。"""
    total_duration = sum(s.duration for s in scenes)
    minutes, seconds = divmod(total_duration, 60)

    animal_counter = Counter([s.zodiac_animal for s in scenes])
    style_counter = Counter([s.image_style for s in scenes])

    animals_zh = [ANIMAL_NAME_ZH[a] for a in configured_animals]
    print(f"共{len(scenes)}个镜头")
    print(f"涉及生肖：[{ '、'.join(animals_zh) }]")
    print(f"预计总时长{minutes}分钟{seconds}秒")
    print("各生肖镜头分布：" + "、".join(f"{ANIMAL_NAME_ZH[a]}{animal_counter.get(a, 0)}个" for a in configured_animals))
    print("各画面类型分布：" + "、".join(f"{k}{v}个" for k, v in style_counter.items()))


def main() -> None:
    """主流程：读入->拆分->识别->生成->导出。"""
    script_path = Path("./script.txt")
    config_path = Path("./config.json")

    try:
        script_text = read_text_file(script_path)
        config = read_config(config_path)

        # 构建镜头文本片段并把数量控制到目标区间
        segments = build_segments(script_text)
        segments = rebalance_scene_count(segments, 35, 55)

        # 逐条生成结构化分镜
        scenes: List[Scene] = []
        for idx, narration in enumerate(segments, start=1):
            animal = detect_animal(
                narration,
                config["zodiac_animals"],
                config["primary_animal"],
            )
            image_style, style_extra = detect_image_style(narration)
            duration = calculate_duration(narration, image_style)
            visual_prompt = build_visual_prompt(animal, style_extra, config["video_theme"])

            scenes.append(
                Scene(
                    scene_id=f"{idx:03d}",
                    narration=narration,
                    visual_prompt=visual_prompt,
                    image_style=image_style,
                    zodiac_animal=animal,
                    duration=duration,
                    motion=MOTION_CYCLE[(idx - 1) % len(MOTION_CYCLE)],
                )
            )

        output_dir = make_output_dir(Path("./project"))
        json_path = output_dir / "storyboard.json"
        xlsx_path = output_dir / "storyboard.xlsx"

        json_data = [asdict(scene) for scene in scenes]
        json_path.write_text(json.dumps(json_data, ensure_ascii=False, indent=2), encoding="utf-8")
        scenes_to_excel(scenes, xlsx_path)

        print(f"输出目录：{output_dir}")
        print(f"已生成：{json_path}")
        print(f"已生成：{xlsx_path}")
        print_stats(scenes, config["zodiac_animals"])

    except SceneSplitterError as exc:
        print(f"[错误] {exc}")
        sys.exit(1)
    except Exception as exc:  # 兜底异常，避免用户看见难懂堆栈
        print(f"[错误] 脚本执行失败：{exc}")
        sys.exit(1)


if __name__ == "__main__":
    main()
