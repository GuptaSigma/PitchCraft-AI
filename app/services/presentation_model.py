"""Shared presentation layout model and helper utilities.

This module normalizes raw presentation JSON into a reusable intermediate
structure that can be rendered by multiple exporters (PPTX, PDF, DOC/DOCX).
The goal is to keep behavioural parity across formats while avoiding the
need to duplicate layout logic in each exporter.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any, Dict, Iterable, List, Optional

# THEME DEFINITIONS (shared across all exporters)
THEMES: Dict[str, Dict[str, Any]] = {
    "dialogue": {
        "name": "Dialogue White",
        "bg": (255, 255, 255),
        "text": (15, 23, 42),
        "accent": (99, 102, 241),
        "card": (248, 250, 252),
    },
    "alien": {
        "name": "Alien Dark",
        "bg": (15, 23, 42),
        "text": (241, 245, 249),
        "accent": (34, 211, 238),
        "card": (30, 41, 59),
    },
    "wine": {
        "name": "Wine Elegance",
        "bg": (88, 28, 60),
        "text": (255, 222, 200),
        "accent": (244, 114, 182),
        "card": (45, 11, 30),
    },
    "snowball": {
        "name": "Snowball Blue",
        "bg": (224, 242, 254),
        "text": (30, 58, 138),
        "accent": (14, 165, 233),
        "card": (186, 230, 253),
    },
    "petrol": {
        "name": "Petrol Steel",
        "bg": (71, 85, 105),
        "text": (241, 245, 249),
        "accent": (14, 165, 233),
        "card": (51, 65, 85),
    },
    "piano": {
        "name": "Piano Contrast",
        "bg": (255, 255, 255),
        "text": (0, 0, 0),
        "accent": (0, 0, 0),
        "card": (245, 245, 245),
    },
    "sunset": {
        "name": "Sunset Orange",
        "bg": (254, 252, 232),
        "text": (120, 53, 15),
        "accent": (249, 115, 22),
        "card": (254, 243, 199),
    },
    "midnight": {
        "name": "Midnight Purple",
        "bg": (30, 41, 59),
        "text": (226, 232, 240),
        "accent": (168, 85, 247),
        "card": (15, 23, 42),
    },
}

# Helper utilities

def calculate_brightness(rgb: Iterable[int]) -> float:
    """Return perceived brightness for an RGB tuple (0-255 values)."""

    r, g, b = rgb
    return (r * 299 + g * 587 + b * 114) / 1000

def readable_text_rgb(background_rgb: Iterable[int]) -> tuple[int, int, int]:
    """Choose black or white text based on the provided background RGB."""

    brightness = calculate_brightness(background_rgb)
    return (0, 0, 0) if brightness > 128 else (255, 255, 255)

def detect_theme(presentation_data: Any, default: str = "dialogue") -> str:
    """Best-effort theme detection shared by all exporters."""

    theme_name: Optional[str] = None

    # Object attribute access
    if hasattr(presentation_data, "theme"):
        theme_name = getattr(presentation_data, "theme")

    # __dict__ access for simple namespaces
    if theme_name is None and hasattr(presentation_data, "__dict__"):
        theme_name = presentation_data.__dict__.get("theme")

    # Dictionary access
    if theme_name is None and isinstance(presentation_data, dict):
        theme_name = presentation_data.get("theme")

    if not theme_name or not str(theme_name).strip():
        theme_name = default

    theme_name = str(theme_name).lower().strip()
    if theme_name not in THEMES:
        theme_name = default
    return theme_name

def parse_content(content: Any) -> List[str]:
    """Normalize arbitrary content values to a list of strings."""

    if content is None:
        return []
    if isinstance(content, list):
        return [str(item).strip() for item in content if str(item).strip()]
    if not isinstance(content, str):
        content = str(content)

    lines = [line.strip() for line in content.split("\n") if line.strip()]
    cleaned: List[str] = []
    for line in lines:
        cleaned_line = line.lstrip("-*•0123456789. \t")
        if cleaned_line:
            cleaned.append(cleaned_line)
    return cleaned

def text_scale(texts: Any, minimum: int = 120, maximum: int = 600) -> float:
    """Return a 0..1 scale based on combined text length."""

    if not texts:
        return 0.0
    if isinstance(texts, str):
        total = len(texts)
    else:
        total = sum(len(str(t)) for t in texts)

    if maximum <= minimum:
        return 1.0
    scale = (total - minimum) / (maximum - minimum)
    return max(0.0, min(1.0, scale))

def scaled_font(scale: float, maximum: int, minimum: int) -> int:
    """Return an integer font size scaled between min/max values."""

    return int(round(maximum - scale * (maximum - minimum)))

def scaled_height(scale: float, minimum: float, maximum: float) -> float:
    """Scale a height value (in inches) based on text length."""

    return minimum + (maximum - minimum) * scale

# Intermediate representation (Blocks / Sections)

@dataclass(slots=True)
class CardBlock:
    title: str
    body: List[str] = field(default_factory=list)
    icon: Optional[str] = None
    image: Optional[str] = None
    extra: Dict[str, Any] = field(default_factory=dict)

@dataclass(slots=True)
class SectionBlock:
    heading: Optional[str]
    description: List[str]
    meta: Dict[str, Any] = field(default_factory=dict)

@dataclass(slots=True)
class SlideBlock:
    index: int
    layout: str
    title: Optional[str]
    subtitle: Optional[str]
    body: List[str]
    image: Optional[str]
    image_alt: Optional[str]
    cards: List[CardBlock] = field(default_factory=list)
    sections: List[SectionBlock] = field(default_factory=list)
    metadata: Dict[str, Any] = field(default_factory=dict)
    raw: Dict[str, Any] = field(default_factory=dict)

    def to_legacy_dict(self) -> Dict[str, Any]:
        """Return a dict-compatible view for legacy layout functions."""

        data = dict(self.raw)
        data.setdefault("layout", self.layout)
        data.setdefault("title", self.title)
        if self.body and "content" not in data:
            data["content"] = self.body
        if self.image and "image" not in data:
            data["image"] = self.image
        if self.cards and "cards" not in data:
            data["cards"] = [card.__dict__ for card in self.cards]
        return data

@dataclass(slots=True)
class PresentationDocument:
    title: str
    theme_key: str
    slides: List[SlideBlock]
    metadata: Dict[str, Any] = field(default_factory=dict)
    raw: Any = None

    @property
    def theme(self) -> Dict[str, Any]:
        return THEMES[self.theme_key]

# Compiler – converts raw presentation data → PresentationDocument

def compile_presentation(presentation_data: Any) -> PresentationDocument:
    """Normalize arbitrary presentation payload into a PresentationDocument."""

    theme_key = detect_theme(presentation_data)

    title = None
    raw_content: Any = None

    if isinstance(presentation_data, dict):
        title = presentation_data.get("title") or presentation_data.get("prompt")
        raw_content = presentation_data.get("content")
    else:
        title = getattr(presentation_data, "title", None) or getattr(
            presentation_data, "prompt", None
        )
        raw_content = getattr(presentation_data, "content", None)

    if isinstance(raw_content, str):
        try:
            import json

            raw_content = json.loads(raw_content)
        except Exception:
            raw_content = None

    slides_payload: List[Dict[str, Any]] = []
    if isinstance(raw_content, dict):
        slides_payload = raw_content.get("slides", []) or []
    elif isinstance(raw_content, list):
        slides_payload = raw_content

    slides: List[SlideBlock] = []
    for index, raw_slide in enumerate(slides_payload, start=1):
        slides.append(_normalize_slide(raw_slide, index))

    return PresentationDocument(
        title=title or "Untitled",
        theme_key=theme_key,
        slides=slides,
        metadata={"slide_count": len(slides)},
        raw=presentation_data,
    )

def _normalize_slide(raw_slide: Any, index: int) -> SlideBlock:
    if not isinstance(raw_slide, dict):
        raw_slide = {"content": raw_slide}

    layout = raw_slide.get("layout", "standard")
    title = raw_slide.get("title")
    subtitle = raw_slide.get("subtitle") or raw_slide.get("subheading")
    body = parse_content(raw_slide.get("content"))
    image = raw_slide.get("image") or raw_slide.get("cover")
    image_alt = raw_slide.get("image_caption") or raw_slide.get("imageAlt")

    cards_data = raw_slide.get("cards") or []
    cards: List[CardBlock] = []
    for card in cards_data:
        if not isinstance(card, dict):
            continue
        cards.append(
            CardBlock(
                title=str(card.get("title") or card.get("heading") or "").strip(),
                body=parse_content(card.get("content") or card.get("body")),
                icon=card.get("icon"),
                image=card.get("image"),
                extra={k: v for k, v in card.items() if k not in {"title", "heading", "content", "body", "icon", "image"}},
            )
        )

    sections_data = raw_slide.get("sections") or raw_slide.get("timeline") or []
    sections: List[SectionBlock] = []
    for section in sections_data:
        if not isinstance(section, dict):
            description = parse_content(section)
            if not description:
                continue
            sections.append(SectionBlock(heading=None, description=description))
            continue
        heading = section.get("title") or section.get("heading")
        description = parse_content(section.get("content") or section.get("description"))
        sections.append(
            SectionBlock(
                heading=str(heading).strip() if heading else None,
                description=description,
                meta={
                    k: v
                    for k, v in section.items()
                    if k
                    not in {
                        "title",
                        "heading",
                        "content",
                        "description",
                    }
                },
            )
        )

    metadata = {
        k: v
        for k, v in raw_slide.items()
        if k
        not in {
            "layout",
            "title",
            "subtitle",
            "subheading",
            "content",
            "image",
            "cover",
            "image_caption",
            "imageAlt",
            "cards",
            "sections",
            "timeline",
        }
    }

    return SlideBlock(
        index=index,
        layout=layout,
        title=str(title).strip() if title else None,
        subtitle=str(subtitle).strip() if subtitle else None,
        body=body,
        image=image,
        image_alt=image_alt,
        cards=cards,
        sections=sections,
        metadata=metadata,
        raw=raw_slide,
    )
