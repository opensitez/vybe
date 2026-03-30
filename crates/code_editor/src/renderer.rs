use std::collections::{HashMap, HashSet};
use std::time::{Duration, Instant};
use std::num::NonZeroU32;
use std::sync::Arc;
use std::fs;
use cosmic_text::{Attrs, Buffer, Color, Family, FontSystem, Metrics, Motion, Shaping, SwashCache, Action, Edit, AttrsList, Cursor, Selection};
use lsp_types::Diagnostic;
use fuzzy_matcher::FuzzyMatcher;
use fuzzy_matcher::skim::SkimMatcherV2;
use tiny_skia::{Color as SkiaColor, Paint, Pixmap, PixmapPaint, Rect, Transform, ColorU8, Stroke, PathBuilder};
use winit::application::ApplicationHandler;
use winit::event::{WindowEvent, ElementState, MouseButton, MouseScrollDelta};
use winit::event_loop::{ControlFlow, EventLoop};
use winit::window::{Window, WindowAttributes};
use winit::keyboard::{Key, NamedKey};
#[cfg(target_os = "macos")]
use winit::platform::modifier_supplement::KeyEventExtModifierSupplement;
use softbuffer::{Context, Surface};
use arboard::Clipboard;

use std::path::Path;
use serde::{Deserialize, Serialize};

use crate::editor::{Editor as MyEditor, TokenKind};
use crate::language::{load_language, LanguageDef};
use crate::lsp_client::{LspClient, LspRequest, LspEvent};
use osz_widgets::{TreeView, TreeEvent, Dropdown, DropdownEvent};

#[derive(Debug, Clone, Deserialize, Serialize)]
pub struct Keybinding {
    pub key: String,
    #[serde(default)] pub cmd: bool,
    #[serde(default)] pub shift: bool,
    #[serde(default)] pub alt: bool,
    pub action: String,
}

const SCALE: f32 = 2.0;
const EXPLORER_WIDTH: f32 = 250.0; 
const TAB_BAR_HEIGHT: f32 = 36.0;
const MINIMAP_WIDTH: f32 = 80.0;
const UI_BAR_HEIGHT: f32 = 0.0;
const FOOTER_HEIGHT: f32 = 24.0;
const GUTTER_WIDTH: f32 = 64.0;
const SPLITTER_WIDTH: f32 = 4.0;
pub struct Tab {
    pub name: String,
    pub path: Option<String>,
    pub widget: CodeEditorWidget,
    pub is_sticky: bool,
    pub buffer: Option<Buffer>,
    pub is_modified: bool,
}

#[derive(Clone, Copy)]
pub struct Theme {
    pub bg: Color,
    pub sidebar_bg: Color,
    pub sidebar_text: Color,
    pub current_line: Color,
    pub selection: Color,
    pub match_highlight: Color,
    pub text: Color,
    pub kw: Color,
    pub type_kw: Color,
    pub comment: Color,
    pub string: Color,
    pub number: Color,
    pub guide: Color,
    pub bracket: Color,
    pub punctuation: Color,
    pub diagnostic_error: Color,
    pub tab_bar_bg: Color,
    pub active_tab_bg: Color,
    pub inactive_tab_bg: Color,
    pub footer_bg: Color,
    pub footer_text: Color,
    pub splitter_bg: Color,
    pub minimap_bg: Color,
    pub gutter_divider: Color,
    pub active_tab_text: Color,
    pub inactive_tab_text: Color,
}

impl Theme {
    pub fn silicon_green() -> Self {
        Self {
            bg: Color::rgb(0x0d, 0x20, 0x16),
            sidebar_bg: Color::rgb(0x08, 0x14, 0x0e),
            sidebar_text: Color::rgb(0x00, 0xff, 0x9f),
            current_line: Color::rgb(0x13, 0x2d, 0x20),
            selection: Color::rgb(0x1a, 0x4d, 0x36),
            match_highlight: Color::rgb(0x00, 0xff, 0x9f),
            text: Color::rgb(0xee, 0xff, 0xf2),
            kw: Color::rgb(0x00, 0xff, 0x9f),
            type_kw: Color::rgb(0x00, 0xe5, 0xff),
            comment: Color::rgb(0x4a, 0x6e, 0x5a),
            string: Color::rgb(0x7f, 0xff, 0x00),
            number: Color::rgb(0x5f, 0xff, 0xd7),
            guide: Color::rgb(0x1a, 0x4d, 0x36),
            bracket: Color::rgb(0xee, 0xff, 0xf2),
            punctuation: Color::rgb(0x00, 0xff, 0x9f),
            diagnostic_error: Color::rgb(0xff, 0x2d, 0x55),
            tab_bar_bg: Color::rgb(0x08, 0x14, 0x0e),
            active_tab_bg: Color::rgb(0x0d, 0x20, 0x16),
            inactive_tab_bg: Color::rgb(0x08, 0x14, 0x0e),
            footer_bg: Color::rgb(0x00, 0xff, 0x9f),
            footer_text: Color::rgb(0x0d, 0x20, 0x16),
            splitter_bg: Color::rgb(0x0d, 0x20, 0x16),
            minimap_bg: Color::rgb(0x0d, 0x20, 0x16),
            gutter_divider: Color::rgb(0x00, 0xff, 0x9f),
            active_tab_text: Color::rgb(0x00, 0xff, 0x9f),
            inactive_tab_text: Color::rgb(0x00, 0x8f, 0x5f),
        }
    }

    pub fn cloud_blue() -> Self {
        Self {
            bg: Color::rgb(0xe0, 0xf2, 0xfe),
            sidebar_bg: Color::rgb(0xba, 0xe6, 0xfd),
            sidebar_text: Color::rgb(0x03, 0x69, 0xa1),
            current_line: Color::rgb(0xf0, 0xf9, 0xff),
            selection: Color::rgb(0x7d, 0xd3, 0xfc),
            match_highlight: Color::rgb(0x02, 0x84, 0xc7),
            text: Color::rgb(0x0c, 0x4a, 0x6e),
            kw: Color::rgb(0x02, 0x84, 0xc7),
            type_kw: Color::rgb(0x03, 0x69, 0xa1),
            comment: Color::rgb(0x64, 0x74, 0x8b),
            string: Color::rgb(0x05, 0x96, 0x69),
            number: Color::rgb(0xdb, 0x27, 0x77),
            guide: Color::rgb(0xba, 0xe6, 0xfd),
            bracket: Color::rgb(0x0c, 0x4a, 0x6e),
            punctuation: Color::rgb(0x0c, 0x4a, 0x6e),
            diagnostic_error: Color::rgb(0xdc, 0x26, 0x26),
            tab_bar_bg: Color::rgb(0xba, 0xe6, 0xfd),
            active_tab_bg: Color::rgb(0xe0, 0xf2, 0xfe),
            inactive_tab_bg: Color::rgb(0xba, 0xe6, 0xfd),
            footer_bg: Color::rgb(0x02, 0x84, 0xc7),
            footer_text: Color::rgb(0xff, 0xff, 0xff),
            splitter_bg: Color::rgb(0xba, 0xe6, 0xfd),
            minimap_bg: Color::rgb(0xe0, 0xf2, 0xfe),
            gutter_divider: Color::rgb(0x03, 0x69, 0xa1),
            active_tab_text: Color::rgb(0x03, 0x69, 0xa1),
            inactive_tab_text: Color::rgb(0x0c, 0x4a, 0x6e),
        }
    }

    pub fn coffee_cream() -> Self {
        Self {
            bg: Color::rgb(0xf5, 0xf5, 0xf4),
            sidebar_bg: Color::rgb(0xe7, 0xe5, 0xe4),
            sidebar_text: Color::rgb(0x78, 0x35, 0x0f),
            current_line: Color::rgb(0xfa, 0xfa, 0xf9),
            selection: Color::rgb(0xd6, 0xd3, 0xd1),
            match_highlight: Color::rgb(0x92, 0x40, 0x0e),
            text: Color::rgb(0x44, 0x40, 0x3c),
            kw: Color::rgb(0x92, 0x40, 0x0e),
            type_kw: Color::rgb(0x78, 0x35, 0x0f),
            comment: Color::rgb(0x78, 0x71, 0x6c),
            string: Color::rgb(0x16, 0x65, 0x34),
            number: Color::rgb(0x9f, 0x12, 0x39),
            guide: Color::rgb(0xd6, 0xd3, 0xd1),
            bracket: Color::rgb(0x44, 0x40, 0x3c),
            punctuation: Color::rgb(0x44, 0x40, 0x3c),
            diagnostic_error: Color::rgb(0xb9, 0x1c, 0x1c),
            tab_bar_bg: Color::rgb(0xe7, 0xe5, 0xe4),
            active_tab_bg: Color::rgb(0xf5, 0xf5, 0xf4),
            inactive_tab_bg: Color::rgb(0xe7, 0xe5, 0xe4),
            footer_bg: Color::rgb(0x78, 0x35, 0x0f),
            footer_text: Color::rgb(0xff, 0xff, 0xff),
            splitter_bg: Color::rgb(0xe7, 0xe5, 0xe4),
            minimap_bg: Color::rgb(0xf5, 0xf5, 0xf4),
            gutter_divider: Color::rgb(0x78, 0x35, 0x0f),
            active_tab_text: Color::rgb(0x78, 0x35, 0x0f),
            inactive_tab_text: Color::rgb(0x44, 0x40, 0x3c),
        }
    }

    pub fn sakura_pink() -> Self {
        Self {
            bg: Color::rgb(0xff, 0xf1, 0xf2),
            sidebar_bg: Color::rgb(0xff, 0xe4, 0xe6),
            sidebar_text: Color::rgb(0x9d, 0x17, 0x4d),
            current_line: Color::rgb(0xff, 0xfb, 0xfc),
            selection: Color::rgb(0xfe, 0xcd, 0xd3),
            match_highlight: Color::rgb(0xdb, 0x27, 0x77),
            text: Color::rgb(0x88, 0x13, 0x37),
            kw: Color::rgb(0xe1, 0x1d, 0x48),
            type_kw: Color::rgb(0xbe, 0x18, 0x5d),
            comment: Color::rgb(0x9d, 0x17, 0x4d),
            string: Color::rgb(0x0d, 0x94, 0x88),
            number: Color::rgb(0xf4, 0x3f, 0x5e),
            guide: Color::rgb(0xfe, 0xcd, 0xd3),
            bracket: Color::rgb(0x88, 0x13, 0x37),
            punctuation: Color::rgb(0xdb, 0x27, 0x77),
            diagnostic_error: Color::rgb(0xbe, 0x12, 0x3c),
            tab_bar_bg: Color::rgb(0xff, 0xe4, 0xe6),
            active_tab_bg: Color::rgb(0xff, 0xf1, 0xf2),
            inactive_tab_bg: Color::rgb(0xff, 0xe4, 0xe6),
            footer_bg: Color::rgb(0xbe, 0x12, 0x3c),
            footer_text: Color::rgb(0xff, 0xff, 0xff),
            splitter_bg: Color::rgb(0xff, 0xe4, 0xe6),
            minimap_bg: Color::rgb(0xff, 0xf1, 0xf2),
            gutter_divider: Color::rgb(0xbe, 0x12, 0x3c),
            active_tab_text: Color::rgb(0x9d, 0x17, 0x4d),
            inactive_tab_text: Color::rgb(0x88, 0x13, 0x37),
        }
    }

    pub fn one_dark() -> Self {
        Self {
            bg: Color::rgb(0x28, 0x2c, 0x34),
            sidebar_bg: Color::rgb(0x21, 0x25, 0x2b),
            sidebar_text: Color::rgb(0xab, 0xb2, 0xbf),
            current_line: Color::rgb(0x2c, 0x31, 0x3c),
            selection: Color::rgb(0x3e, 0x44, 0x51),
            match_highlight: Color::rgb(0x48, 0x4e, 0x5b),
            text: Color::rgb(0xab, 0xb2, 0xbf),
            kw: Color::rgb(0xc6, 0x78, 0xdd),
            type_kw: Color::rgb(0x61, 0xaf, 0xef),
            comment: Color::rgb(0x5c, 0x63, 0x70),
            string: Color::rgb(0x98, 0xc3, 0x79),
            number: Color::rgb(0xd1, 0x9a, 0x66),
            guide: Color::rgb(0x4b, 0x52, 0x63),
            bracket: Color::rgb(0xab, 0xb2, 0xbf),
            punctuation: Color::rgb(0xab, 0xb2, 0xbf),
            diagnostic_error: Color::rgb(0xe0, 0x6c, 0x75),
            tab_bar_bg: Color::rgb(0x21, 0x25, 0x2b),
            active_tab_bg: Color::rgb(0x28, 0x2c, 0x34),
            inactive_tab_bg: Color::rgb(0x21, 0x25, 0x2b),
            footer_bg: Color::rgb(0x21, 0x25, 0x2b),
            footer_text: Color::rgb(0xab, 0xb2, 0xbf),
            splitter_bg: Color::rgb(0x18, 0x1a, 0x1f),
            minimap_bg: Color::rgb(0x1a, 0x1a, 0x20),
            gutter_divider: Color::rgb(0x3b, 0x40, 0x48),
            active_tab_text: Color::rgb(0xff, 0xff, 0xff),
            inactive_tab_text: Color::rgb(0xab, 0xb2, 0xbf),
        }
    }

    pub fn monokai() -> Self {
        Self {
            bg: Color::rgb(0x27, 0x28, 0x22),
            sidebar_bg: Color::rgb(0x1e, 0x1f, 0x1c),
            sidebar_text: Color::rgb(0xf8, 0xf8, 0xf2),
            current_line: Color::rgb(0x3e, 0x3d, 0x32),
            selection: Color::rgb(0x49, 0x48, 0x3e),
            match_highlight: Color::rgb(0xa6, 0xe2, 0x2e),
            text: Color::rgb(0xf8, 0xf8, 0xf2),
            kw: Color::rgb(0xf9, 0x26, 0x72),
            type_kw: Color::rgb(0x66, 0xd9, 0xef),
            comment: Color::rgb(0x75, 0x71, 0x5e),
            string: Color::rgb(0xe6, 0xdb, 0x74),
            number: Color::rgb(0xae, 0x81, 0xff),
            guide: Color::rgb(0x49, 0x48, 0x3e),
            bracket: Color::rgb(0xf8, 0xf8, 0xf2),
            punctuation: Color::rgb(0xf8, 0xf8, 0xf2),
            diagnostic_error: Color::rgb(0xf9, 0x26, 0x72),
            tab_bar_bg: Color::rgb(0x1e, 0x1f, 0x1c),
            active_tab_bg: Color::rgb(0x27, 0x28, 0x22),
            inactive_tab_bg: Color::rgb(0x1e, 0x1f, 0x1c),
            footer_bg: Color::rgb(0x1e, 0x1f, 0x1c),
            footer_text: Color::rgb(0xf8, 0xf8, 0xf2),
            splitter_bg: Color::rgb(0x1e, 0x1f, 0x1c),
            minimap_bg: Color::rgb(0x16, 0x16, 0x16),
            gutter_divider: Color::rgb(0x49, 0x48, 0x3e),
            active_tab_text: Color::rgb(0xff, 0xff, 0xff),
            inactive_tab_text: Color::rgb(0xf8, 0xf8, 0xf2),
        }
    }

    pub fn frost_light() -> Self {
        Self {
            bg: Color::rgb(0xf2, 0xf7, 0xff),
            sidebar_bg: Color::rgb(0xe8, 0xf0, 0xfa),
            sidebar_text: Color::rgb(0x47, 0x55, 0x69),
            current_line: Color::rgb(0xff, 0xff, 0xff),
            selection: Color::rgb(0xbf, 0xdb, 0xfe),
            match_highlight: Color::rgb(0x3b, 0x82, 0xf6),
            text: Color::rgb(0x1e, 0x29, 0x3b),
            kw: Color::rgb(0x25, 0x63, 0xeb),
            type_kw: Color::rgb(0x08, 0x91, 0xb2),
            comment: Color::rgb(0x94, 0xa3, 0xb8),
            string: Color::rgb(0x05, 0x96, 0x69),
            number: Color::rgb(0x7c, 0x3a, 0xed),
            guide: Color::rgb(0xe2, 0xe8, 0xf0),
            bracket: Color::rgb(0x1e, 0x29, 0x3b),
            punctuation: Color::rgb(0x64, 0x74, 0x8b),
            diagnostic_error: Color::rgb(0xe1, 0x1d, 0x48),
            tab_bar_bg: Color::rgb(0xe8, 0xf0, 0xfa),
            active_tab_bg: Color::rgb(0xf2, 0xf7, 0xff),
            inactive_tab_bg: Color::rgb(0xe8, 0xf0, 0xfa),
            footer_bg: Color::rgb(0x3b, 0x82, 0xf6),
            footer_text: Color::rgb(0xff, 0xff, 0xff),
            splitter_bg: Color::rgb(0xe2, 0xe8, 0xf0),
            minimap_bg: Color::rgb(0xf8, 0xfa, 0xfc),
            gutter_divider: Color::rgb(0x3b, 0x82, 0xf6),
            active_tab_text: Color::rgb(0x1e, 0x29, 0x3b),
            inactive_tab_text: Color::rgb(0x64, 0x74, 0x8b),
        }
    }

    pub fn solarized_light() -> Self {
        Self {
            bg: Color::rgb(0xfd, 0xf6, 0xe3),
            sidebar_bg: Color::rgb(0xee, 0xe8, 0xd5),
            sidebar_text: Color::rgb(0x58, 0x6e, 0x75),
            current_line: Color::rgb(0xee, 0xe8, 0xd5),
            selection: Color::rgb(0x26, 0x8b, 0xd2),
            match_highlight: Color::rgb(0x85, 0x99, 0x00),
            text: Color::rgb(0x65, 0x7b, 0x83),
            kw: Color::rgb(0x85, 0x99, 0x00),
            type_kw: Color::rgb(0xb5, 0x89, 0x00),
            comment: Color::rgb(0x93, 0xa1, 0xa1),
            string: Color::rgb(0x2a, 0xa1, 0x98),
            number: Color::rgb(0xd3, 0x36, 0x82),
            guide: Color::rgb(0x93, 0xa1, 0xa1),
            bracket: Color::rgb(0x65, 0x7b, 0x83),
            punctuation: Color::rgb(0x65, 0x7b, 0x83),
            diagnostic_error: Color::rgb(0xdc, 0x32, 0x2f),
            tab_bar_bg: Color::rgb(0xee, 0xe8, 0xd5),
            active_tab_bg: Color::rgb(0xfd, 0xf6, 0xe3),
            inactive_tab_bg: Color::rgb(0xee, 0xe8, 0xd5),
            footer_bg: Color::rgb(0x07, 0x36, 0x42),
            footer_text: Color::rgb(0xee, 0xe8, 0xd5),
            splitter_bg: Color::rgb(0xee, 0xe8, 0xd5),
            minimap_bg: Color::rgb(0xee, 0xe8, 0xd5),
            gutter_divider: Color::rgb(0x93, 0xa1, 0xa1),
            active_tab_text: Color::rgb(0x58, 0x6e, 0x75),
            inactive_tab_text: Color::rgb(0x93, 0xa1, 0xa1),
        }
    }

    pub fn midnight() -> Self {
        Self {
            bg: Color::rgb(0x05, 0x05, 0x05),
            sidebar_bg: Color::rgb(0x00, 0x00, 0x00),
            sidebar_text: Color::rgb(0xff, 0xff, 0xff),
            current_line: Color::rgb(0x1a, 0x1a, 0x1a),
            selection: Color::rgb(0x33, 0x33, 0x33),
            match_highlight: Color::rgb(0xff, 0xff, 0xff),
            text: Color::rgb(0xff, 0xff, 0xff),
            kw: Color::rgb(0x00, 0xf6, 0xff),
            type_kw: Color::rgb(0xbd, 0x93, 0xf9),
            comment: Color::rgb(0x62, 0x72, 0xa4),
            string: Color::rgb(0x50, 0xfa, 0x7b),
            number: Color::rgb(0xff, 0xb8, 0x6c),
            guide: Color::rgb(0x33, 0x33, 0x33),
            bracket: Color::rgb(0xff, 0xff, 0xff),
            punctuation: Color::rgb(0xff, 0xff, 0xff),
            diagnostic_error: Color::rgb(0xff, 0x55, 0x55),
            tab_bar_bg: Color::rgb(0x00, 0x00, 0x00),
            active_tab_bg: Color::rgb(0x05, 0x05, 0x05),
            inactive_tab_bg: Color::rgb(0x00, 0x00, 0x00),
            footer_bg: Color::rgb(0x00, 0x00, 0x00),
            footer_text: Color::rgb(0xff, 0xff, 0xff),
            splitter_bg: Color::rgb(0x05, 0x05, 0x05),
            minimap_bg: Color::rgb(0x00, 0x00, 0x00),
            gutter_divider: Color::rgb(0x33, 0x33, 0x33),
            active_tab_text: Color::rgb(0xff, 0xff, 0xff),
            inactive_tab_text: Color::rgb(0x88, 0x88, 0x88),
        }
    }

    pub fn aura() -> Self {
        Self {
            bg: Color::rgb(0x15, 0x14, 0x1b),
            sidebar_bg: Color::rgb(0x1a, 0x19, 0x22),
            sidebar_text: Color::rgb(0xde, 0xde, 0xde),
            current_line: Color::rgb(0x1c, 0x1b, 0x24),
            selection: Color::rgb(0x3d, 0x37, 0x5e),
            match_highlight: Color::rgb(0xa2, 0x77, 0xff),
            text: Color::rgb(0xed, 0xe0, 0xeb),
            kw: Color::rgb(0xa2, 0x77, 0xff),
            type_kw: Color::rgb(0xff, 0xca, 0x85),
            comment: Color::rgb(0x61, 0x61, 0x61),
            string: Color::rgb(0x61, 0xff, 0xca),
            number: Color::rgb(0xff, 0x67, 0x67),
            guide: Color::rgb(0x3b, 0x33, 0x4b),
            bracket: Color::rgb(0xed, 0xe0, 0xeb),
            punctuation: Color::rgb(0xed, 0xe0, 0xeb),
            diagnostic_error: Color::rgb(0xff, 0x67, 0x67),
            tab_bar_bg: Color::rgb(0x15, 0x14, 0x1b),
            active_tab_bg: Color::rgb(0x15, 0x14, 0x1b),
            inactive_tab_bg: Color::rgb(0x1a, 0x19, 0x22),
            footer_bg: Color::rgb(0x1b, 0x1a, 0x23),
            footer_text: Color::rgb(0xed, 0xe0, 0xeb),
            splitter_bg: Color::rgb(0x1c, 0x1b, 0x24),
            minimap_bg: Color::rgb(0x10, 0x10, 0x10),
            gutter_divider: Color::rgb(0x3d, 0x37, 0x5e),
            active_tab_text: Color::rgb(0xed, 0xe0, 0xeb),
            inactive_tab_text: Color::rgb(0xde, 0xde, 0xde),
        }
    }

    pub fn veridian() -> Self {
        Self {
            bg: Color::rgb(0x0c, 0x20, 0x1d),
            sidebar_bg: Color::rgb(0x08, 0x16, 0x14),
            sidebar_text: Color::rgb(0x40, 0xc0, 0xab),
            current_line: Color::rgb(0x12, 0x2d, 0x2a),
            selection: Color::rgb(0x1a, 0x4d, 0x44),
            match_highlight: Color::rgb(0x00, 0xf5, 0xb1),
            text: Color::rgb(0xe0, 0xf0, 0xee),
            kw: Color::rgb(0x00, 0xf5, 0xb1),
            type_kw: Color::rgb(0x75, 0xff, 0xda),
            comment: Color::rgb(0x4a, 0x6e, 0x66),
            string: Color::rgb(0x95, 0xff, 0x8a),
            number: Color::rgb(0x5f, 0xff, 0xd7),
            guide: Color::rgb(0x1a, 0x4d, 0x44),
            bracket: Color::rgb(0xe0, 0xf0, 0xee),
            punctuation: Color::rgb(0x40, 0xc0, 0xab),
            diagnostic_error: Color::rgb(0xff, 0x2d, 0x55),
            tab_bar_bg: Color::rgb(0x08, 0x16, 0x14),
            active_tab_bg: Color::rgb(0x0c, 0x20, 0x1d),
            inactive_tab_bg: Color::rgb(0x08, 0x16, 0x14),
            footer_bg: Color::rgb(0x00, 0xf5, 0xb1),
            footer_text: Color::rgb(0x0c, 0x20, 0x1d),
            splitter_bg: Color::rgb(0x12, 0x2d, 0x2a),
            minimap_bg: Color::rgb(0x08, 0x16, 0x14),
            gutter_divider: Color::rgb(0x00, 0xf5, 0xb1),
            active_tab_text: Color::rgb(0x00, 0xf5, 0xb1),
            inactive_tab_text: Color::rgb(0x20, 0x60, 0x55),
        }
    }

    pub fn rose() -> Self {
        Self {
            bg: Color::rgb(0x1a, 0x10, 0x12),
            sidebar_bg: Color::rgb(0x14, 0x0c, 0x0d),
            sidebar_text: Color::rgb(0xe0, 0x60, 0x70),
            current_line: Color::rgb(0x28, 0x18, 0x1b),
            selection: Color::rgb(0x4d, 0x20, 0x26),
            match_highlight: Color::rgb(0xe0, 0x60, 0x70),
            text: Color::rgb(0xf8, 0xe8, 0xea),
            kw: Color::rgb(0xe0, 0x60, 0x70),
            type_kw: Color::rgb(0xf0, 0xa0, 0xb0),
            comment: Color::rgb(0x6e, 0x4a, 0x4e),
            string: Color::rgb(0xf0, 0xd0, 0xa0),
            number: Color::rgb(0xf0, 0x80, 0x90),
            guide: Color::rgb(0x4d, 0x1a, 0x21),
            bracket: Color::rgb(0xf8, 0xe8, 0xea),
            punctuation: Color::rgb(0xe0, 0x60, 0x70),
            diagnostic_error: Color::rgb(0xff, 0x20, 0x40),
            tab_bar_bg: Color::rgb(0x14, 0x0c, 0x0d),
            active_tab_bg: Color::rgb(0x1a, 0x10, 0x12),
            inactive_tab_bg: Color::rgb(0x14, 0x0c, 0x0d),
            footer_bg: Color::rgb(0xe0, 0x60, 0x70),
            footer_text: Color::rgb(0x1a, 0x10, 0x12),
            splitter_bg: Color::rgb(0x1a, 0x10, 0x12),
            minimap_bg: Color::rgb(0x14, 0x0c, 0x0d),
            gutter_divider: Color::rgb(0xe0, 0x60, 0x70),
            active_tab_text: Color::rgb(0xe0, 0x60, 0x70),
            inactive_tab_text: Color::rgb(0x70, 0x30, 0x38),
        }
    }

    pub fn cyber() -> Self {
        Self {
            bg: Color::rgb(0x05, 0x05, 0x05),
            sidebar_bg: Color::rgb(0x00, 0x00, 0x00),
            sidebar_text: Color::rgb(0x00, 0xff, 0xff),
            current_line: Color::rgb(0x10, 0x10, 0x10),
            selection: Color::rgb(0x00, 0x40, 0x40),
            match_highlight: Color::rgb(0xff, 0x00, 0xff),
            text: Color::rgb(0xe0, 0xe0, 0xe0),
            kw: Color::rgb(0x00, 0xff, 0xff),
            type_kw: Color::rgb(0xff, 0x00, 0xff),
            comment: Color::rgb(0x40, 0x40, 0x40),
            string: Color::rgb(0xff, 0xff, 0x00),
            number: Color::rgb(0x00, 0xff, 0x00),
            guide: Color::rgb(0x20, 0x20, 0x20),
            bracket: Color::rgb(0x00, 0xff, 0xff),
            punctuation: Color::rgb(0xff, 0x00, 0xff),
            diagnostic_error: Color::rgb(0xff, 0x00, 0x00),
            tab_bar_bg: Color::rgb(0x00, 0x00, 0x00),
            active_tab_bg: Color::rgb(0x05, 0x05, 0x05),
            inactive_tab_bg: Color::rgb(0x00, 0x00, 0x00),
            footer_bg: Color::rgb(0x00, 0xff, 0xff),
            footer_text: Color::rgb(0x00, 0x00, 0x00),
            splitter_bg: Color::rgb(0x05, 0x05, 0x05),
            minimap_bg: Color::rgb(0x00, 0x00, 0x00),
            gutter_divider: Color::rgb(0x00, 0xff, 0xff),
            active_tab_text: Color::rgb(0x00, 0xff, 0xff),
            inactive_tab_text: Color::rgb(0x00, 0x80, 0x80),
        }
    }

    pub fn titanium() -> Self {
        Self {
            bg: Color::rgb(0x1e, 0x1e, 0x22),
            sidebar_bg: Color::rgb(0x16, 0x16, 0x1a),
            sidebar_text: Color::rgb(0xd0, 0xd0, 0xe0),
            current_line: Color::rgb(0x25, 0x25, 0x2d),
            selection: Color::rgb(0x35, 0x35, 0x45),
            match_highlight: Color::rgb(0xa0, 0xa0, 0xf0),
            text: Color::rgb(0xe0, 0xe0, 0xf0),
            kw: Color::rgb(0xa0, 0xa0, 0xf0),
            type_kw: Color::rgb(0xc0, 0xc0, 0xff),
            comment: Color::rgb(0x60, 0x60, 0x75),
            string: Color::rgb(0xa0, 0xd0, 0xa0),
            number: Color::rgb(0xf0, 0x90, 0x90),
            guide: Color::rgb(0x30, 0x30, 0x40),
            bracket: Color::rgb(0xd0, 0xd0, 0xe0),
            punctuation: Color::rgb(0xd0, 0xd0, 0xe0),
            diagnostic_error: Color::rgb(0xff, 0x40, 0x40),
            tab_bar_bg: Color::rgb(0x16, 0x16, 0x1a),
            active_tab_bg: Color::rgb(0x1e, 0x1e, 0x22),
            inactive_tab_bg: Color::rgb(0x16, 0x16, 0x1a),
            footer_bg: Color::rgb(0xd0, 0xd0, 0xe0),
            footer_text: Color::rgb(0x16, 0x16, 0x1a),
            splitter_bg: Color::rgb(0x1e, 0x1e, 0x22),
            minimap_bg: Color::rgb(0x16, 0x16, 0x1a),
            gutter_divider: Color::rgb(0xd0, 0xd0, 0xe0),
            active_tab_text: Color::rgb(0xd0, 0xd0, 0xe0),
            inactive_tab_text: Color::rgb(0x60, 0x60, 0x70),
        }
    }

    pub fn indigo_night() -> Self {
        Self {
            bg: Color::rgb(0x0d, 0x11, 0x17),
            sidebar_bg: Color::rgb(0x01, 0x04, 0x09),
            sidebar_text: Color::rgb(0x58, 0xa6, 0xff),
            current_line: Color::rgb(0x16, 0x1b, 0x22),
            selection: Color::rgb(0x26, 0x4f, 0x78),
            match_highlight: Color::rgb(0x58, 0xa6, 0xff),
            text: Color::rgb(0xc9, 0xd1, 0xd9),
            kw: Color::rgb(0xff, 0x7b, 0x72),
            type_kw: Color::rgb(0x79, 0xc0, 0xff),
            comment: Color::rgb(0x8b, 0x94, 0x9e),
            string: Color::rgb(0xa5, 0xd6, 0xff),
            number: Color::rgb(0x79, 0xc0, 0xff),
            guide: Color::rgb(0x30, 0x36, 0x3d),
            bracket: Color::rgb(0xc9, 0xd1, 0xd9),
            punctuation: Color::rgb(0xc9, 0xd1, 0xd9),
            diagnostic_error: Color::rgb(0xf8, 0x51, 0x49),
            tab_bar_bg: Color::rgb(0x01, 0x04, 0x09),
            active_tab_bg: Color::rgb(0x0d, 0x11, 0x17),
            inactive_tab_bg: Color::rgb(0x01, 0x04, 0x09),
            footer_bg: Color::rgb(0x58, 0xa6, 0xff),
            footer_text: Color::rgb(0x01, 0x04, 0x09),
            splitter_bg: Color::rgb(0x0d, 0x11, 0x17),
            minimap_bg: Color::rgb(0x0d, 0x11, 0x17),
            gutter_divider: Color::rgb(0x58, 0xa6, 0xff),
            active_tab_text: Color::rgb(0x58, 0xa6, 0xff),
            inactive_tab_text: Color::rgb(0x1f, 0x40, 0x60),
        }
    }

    pub fn dark() -> Self {
        Self::one_dark()
    }
}

struct CachedGlyph {
    pixmap: Pixmap,
    left: i32,
    top: i32,
}

pub fn apply_highlighting(editor: &mut cosmic_text::Editor<'static>, my_editor: &MyEditor, attrs: &Attrs, lang: &LanguageDef, theme: &Theme) {
    editor.with_buffer_mut(|buffer| {
        let mut byte_offset = 0usize;
        for (li, line) in buffer.lines.iter_mut().enumerate() {
            let mut list = AttrsList::new(attrs);
            if li < my_editor.line_tokens.len() {
                for token in &my_editor.line_tokens[li] {
                    let color = match token.kind {
                        TokenKind::LineComment | TokenKind::BlockComment => theme.comment,
                        TokenKind::String => theme.string,
                        TokenKind::Number => theme.number,
                        TokenKind::Punct => theme.punctuation,
                        TokenKind::Identifier => {
                            let text = my_editor.slice(token.start, token.end);
                            if lang.keywords.contains(&text) { theme.kw }
                            else if lang.type_keywords.contains(&text) { theme.type_kw }
                            else if lang.constants.contains(&text) { theme.number }
                            else {
                                // Heuristic: if identifier is followed by '(', highlight as function
                                let next_char = my_editor.rope.chars().nth(my_editor.rope.byte_to_char(token.end));
                                if next_char == Some('(') { theme.kw }
                                else if text.chars().next().map(|c| c.is_uppercase()).unwrap_or(false) { theme.type_kw }
                                else { theme.text }
                            }
                        }
                        _ => theme.text,
                    };
                    let start = token.start.saturating_sub(byte_offset);
                    let end = token.end.saturating_sub(byte_offset);
                    list.add_span(start..end, &attrs.clone().color(color));
                }
            }
            line.set_attrs_list(list);
            if li < my_editor.rope.len_lines() { byte_offset += my_editor.rope.line(li).len_bytes(); }
            else { byte_offset += line.text().len() + 1; }
        }
    });
}

pub struct CodeEditorWidget {
    editor: cosmic_text::Editor<'static>,
    my_editor: MyEditor,
    lang_def: LanguageDef,
    theme: Theme,
    metrics: Metrics,
    glyph_cache: HashMap<(cosmic_text::CacheKey, Color, bool), CachedGlyph>,
    digit_cache: Vec<CachedGlyph>,
    needs_reshape: bool,
    pub scroll_y: f32,
    search_query: String,
    replace_query: String,
    is_search_open: bool,
    is_replace_open: bool,
    pub case_sensitive: bool,
    context_menu: Option<((f32, f32), Vec<String>)>,
    font_size: f32,
    pub show_whitespace: bool,
    pub lsp: Arc<LspClient>,
    minimap_pixmap: Option<Pixmap>,
    minimap_needs_redraw: bool,
    pub wrap_lines: bool,
    pub diagnostics: Vec<Diagnostic>,
    pub matching_bracket: Option<usize>,
}

impl CodeEditorWidget {
    pub fn new(mut my_editor: MyEditor, font_system: &mut FontSystem, lsp: Arc<LspClient>, uri: String) -> Self {
        let font_size = 14.0;
        let metrics = Metrics::new(font_size, 20.0).scale(SCALE);
        let lang_def = load_language("rust").unwrap_or_else(|| LanguageDef {
            keywords: HashSet::new(), type_keywords: HashSet::new(), constants: HashSet::new(), operators: HashSet::new(), ignore_case: false, comments: None, brackets: Vec::new(),
        });
        let theme = Theme::dark();
        my_editor.retokenize_all(&lang_def);
        let mut buffer = Buffer::new(font_system, metrics);
        let text = my_editor.rope.to_string();
        buffer.set_text(font_system, &text, &Attrs::new().family(Family::Monospace), Shaping::Advanced, None);
        
        lsp.send(LspRequest::Init(text, "rust".to_string(), uri));

        let mut widget = Self { editor: cosmic_text::Editor::new(buffer), my_editor, lang_def, theme, metrics, glyph_cache: HashMap::new(), digit_cache: Vec::new(), needs_reshape: true, scroll_y: 0.0, search_query: String::new(), replace_query: String::new(), is_search_open: false, is_replace_open: false, case_sensitive: false, context_menu: None, font_size, show_whitespace: false, lsp, minimap_pixmap: None, minimap_needs_redraw: true, wrap_lines: false, diagnostics: Vec::new(), matching_bracket: None };
        widget.update_digit_cache(font_system);
        widget
    }

    fn update_digit_cache(&mut self, font_system: &mut FontSystem) {
        self.digit_cache.clear();
        let mut swash_cache = SwashCache::new();
        let digit_color = Color::rgb(0x85, 0x85, 0x85);
        for i in 0..10 {
            let mut lab = Buffer::new(font_system, self.metrics);
            lab.set_text(font_system, &format!("{}", i), &Attrs::new().family(Family::Monospace).color(digit_color), Shaping::Advanced, None);
            lab.shape_until_scroll(font_system, false);
            if let Some(r) = lab.layout_runs().next() {
                for g in r.glyphs {
                    let pg = g.physical((0.0, 0.0), 1.0);
                    if let Some(img) = swash_cache.get_image(font_system, pg.cache_key) {
                        let mut p = Pixmap::new(img.placement.width.max(1), img.placement.height.max(1)).unwrap();
                        let (r, g, b, a) = (digit_color.r(), digit_color.g(), digit_color.b(), digit_color.a());
                        for (idx, &alpha) in img.data.iter().enumerate() { let af = (alpha as f32 / 255.0) * (a as f32 / 255.0); p.pixels_mut()[idx] = ColorU8::from_rgba((r as f32 * af) as u8, (g as f32 * af) as u8, (b as f32 * af) as u8, (255.0 * af) as u8).premultiply(); }
                        self.digit_cache.push(CachedGlyph { pixmap: p, left: img.placement.left, top: img.placement.top }); break;
                    }
                }
            }
            if self.digit_cache.len() <= i { self.digit_cache.push(CachedGlyph { pixmap: Pixmap::new(1, 1).unwrap(), left: 0, top: 0 }); }
        }
    }

    pub fn set_zoom(&mut self, fs: &mut FontSystem, delta: f32) {
        self.font_size = (self.font_size + delta).clamp(6.0, 72.0);
        self.metrics = Metrics::new(self.font_size, self.font_size * 1.5).scale(SCALE);
        self.editor.with_buffer_mut(|b| b.set_metrics(fs, self.metrics));
        self.glyph_cache.clear(); self.update_digit_cache(fs); self.needs_reshape = true; self.minimap_needs_redraw = true; self.sync();
    }

    pub fn set_language(&mut self, lang_name: &str, uri: String) {
        if let Some(lang) = load_language(lang_name) {
            self.lang_def = lang; 
            self.my_editor.retokenize_all(&self.lang_def); 
            self.my_editor.diagnostics.clear();
            let text = self.my_editor.rope.to_string();
            self.lsp.send(LspRequest::Init(text, lang_name.to_string(), uri));
            self.reapply_highlighting();
            self.needs_reshape = true;
            self.minimap_needs_redraw = true;
        }
    }

    fn reapply_highlighting(&mut self) {
        let theme = self.theme.clone();
        let lang_def = &self.lang_def;
        self.editor.with_buffer_mut(|buffer| {
            for (li, tokens) in self.my_editor.line_tokens.iter().enumerate() {
                if let Some(line) = buffer.lines.get_mut(li) {
                    let line_text = line.text();
                    let mut attrs_list = AttrsList::new(&Attrs::new().family(Family::Monospace).color(theme.text));
                    for token in tokens {
                        let mut col = theme.text;
                        match token.kind {
                            TokenKind::Identifier => {
                                let start = token.start.min(line_text.len());
                                let end = token.end.min(line_text.len());
                                let word = &line_text[start..end];
                                let mut is_kw = lang_def.keywords.contains(word) || lang_def.type_keywords.contains(word);
                                let mut is_const = lang_def.constants.contains(word);
                                
                                if !is_kw && !is_const && lang_def.ignore_case {
                                    let upper = word.to_uppercase();
                                    is_kw = lang_def.keywords.contains(&upper) || lang_def.type_keywords.contains(&upper);
                                    is_const = lang_def.constants.contains(&upper);
                                }

                                if is_kw { col = theme.kw; }
                                else if is_const { col = theme.string; }
                            }
                            TokenKind::String => col = theme.string,
                            TokenKind::Number => col = theme.number,
                            TokenKind::LineComment | TokenKind::BlockComment => col = theme.comment,
                            _ => {}
                        }
                        let start = token.start.min(line_text.len());
                        let end = token.end.min(line_text.len());
                        attrs_list.add_span(start..end, &Attrs::new().family(Family::Monospace).color(col));
                    }
                    line.set_attrs_list(attrs_list);
                }
            }
        });
    }

    fn get_offsets(&self, rect: Rect) -> (f32, f32) { (rect.left() + GUTTER_WIDTH * SCALE, rect.top()) }

    fn is_line_hidden(&self, li: usize) -> bool {
        for (s, e) in &self.my_editor.folds { if self.my_editor.collapsed_starts.contains(s) && li > *s && li <= *e { return true; } } false
    }


    fn find_next(&mut self, _fs: &mut FontSystem) {
        if self.search_query.is_empty() { return; }
        let cursor = self.editor.cursor();
        let total = self.my_editor.rope.to_string();
        let (query, content) = if self.case_sensitive { (self.search_query.clone(), total) } else { (self.search_query.to_lowercase(), total.to_lowercase()) };
        
        let mut start = 0;
        self.editor.with_buffer(|b| { for l in b.lines.iter().take(cursor.line) { start += l.text().len() + 1; } start += cursor.index; });
        let match_idx = content[start.min(content.len())..].find(&query).map(|i| i + start).or_else(|| content.find(&query));
        if let Some(idx) = match_idx {
            let mut cb: usize = 0; 
            let mut tl: usize = 0; 
            let mut tc: usize = 0;
            for (li, text) in content.split('\n').enumerate() { 
                let l_len = text.len();
                if idx >= cb && idx <= cb + l_len { tl = li; tc = idx - cb; break; } 
                cb += l_len + 1; 
            }
            self.editor.set_cursor(Cursor::new(tl, tc)); 
            self.editor.set_selection(Selection::Normal(Cursor::new(tl, tc + self.search_query.len()))); 
            self.needs_reshape = true;
        }
    }

    pub fn toggle_comment(&mut self, fs: &mut FontSystem) {
        let prefix = match &self.lang_def.comments {
            Some(c) => c.line_comment.as_deref().unwrap_or("//"),
            None => "//",
        };
        let (start, end) = self.editor.selection_bounds().map(|(s, e)| (s.line, e.line)).unwrap_or((self.editor.cursor().line, self.editor.cursor().line));
        let mut lines: Vec<String> = self.editor.with_buffer(|b| b.lines.iter().map(|l| l.text().to_string()).collect());
        let all_commented = (start..=end).all(|li| lines[li].trim().starts_with(prefix));
        for li in start..=end {
            if all_commented { if let Some(idx) = lines[li].find(prefix) { lines[li].replace_range(idx..idx+prefix.len(), ""); } }
            else { let first_char = lines[li].chars().take_while(|c| c.is_whitespace()).count(); lines[li].insert_str(first_char, prefix); }
        }
        let text = lines.join("\n");
        self.editor.with_buffer_mut(|b| b.set_text(fs, &text, &Attrs::new().family(Family::Monospace), Shaping::Advanced, None));
        self.needs_reshape = true; self.sync();
    }

    fn draw_squiggle(&self, pixmap: &mut Pixmap, x: f32, y: f32, w: f32, color: Color) {
        let mut pb = PathBuilder::new();
        let mut cx = x;
        let step = 3.0 * SCALE;
        let amp = 1.5 * SCALE;
        pb.move_to(cx, y);
        while cx < x + w {
            pb.quad_to(cx + step/2.0, y + amp, cx + step, y);
            pb.quad_to(cx + step * 1.5, y - amp, cx + step * 2.0, y);
            cx += step * 2.0;
        }
        if let Some(path) = pb.finish() {
            let mut paint = Paint::default();
            paint.set_color_rgba8(color.r(), color.g(), color.b(), 200);
            pixmap.stroke_path(&path, &paint, &Stroke { width: 1.0 * SCALE, ..Default::default() }, Transform::identity(), None);
        }
    }

    pub fn render(&mut self, pixmap: &mut Pixmap, fs: &mut FontSystem, sc: &mut SwashCache, rect: Rect) {
        if self.needs_reshape { apply_highlighting(&mut self.editor, &self.my_editor, &Attrs::new().family(Family::Monospace), &self.lang_def, &self.theme); self.editor.shape_as_needed(fs, false); self.needs_reshape = false; }
        let (x_off, y_off) = self.get_offsets(rect);
        let mut sp = Paint::default(); sp.set_color_rgba8(self.theme.sidebar_bg.r(), self.theme.sidebar_bg.g(), self.theme.sidebar_bg.b(), self.theme.sidebar_bg.a());
        pixmap.fill_rect(Rect::from_xywh(rect.left(), rect.top(), GUTTER_WIDTH * SCALE, rect.height()).unwrap(), &sp, Transform::identity(), None);
        
        // Git Gutter Indicator: Emerald bar for modified lines (simulated for now)
        let mut gp = Paint::default(); gp.set_color_rgba8(34, 197, 94, 200); // emerald-500
        pixmap.fill_rect(Rect::from_xywh(rect.left() + GUTTER_WIDTH * SCALE - 3.0 * SCALE, rect.top(), 2.0 * SCALE, rect.height()).unwrap(), &gp, Transform::identity(), None);

        let mut sep_p = Paint::default(); sep_p.set_color_rgba8(self.theme.gutter_divider.r(), self.theme.gutter_divider.g(), self.theme.gutter_divider.b(), 255);
        pixmap.fill_rect(Rect::from_xywh(rect.left() + (GUTTER_WIDTH - 2.0) * SCALE, rect.top(), 1.0 * SCALE, rect.height()).unwrap(), &sep_p, Transform::identity(), None);

        // 80-Character Ruler: A subtle vertical guide
        let char_w = 8.4 * (self.metrics.font_size / 14.0) * SCALE;
        let ruler_x = x_off + (80.0 * char_w);
        if ruler_x < rect.right() - MINIMAP_WIDTH * SCALE {
            let mut rp = Paint::default(); rp.set_color_rgba8(255, 255, 255, 20);
            pixmap.fill_rect(Rect::from_xywh(ruler_x, rect.top(), 1.0, rect.height()).unwrap(), &rp, Transform::identity(), None);
        }

        let cursor_state = self.editor.cursor(); let selection = self.editor.selection_bounds();
        let selected_text = self.editor.copy_selection();
        let partner = self.my_editor.find_matching_bracket(cursor_state.line, cursor_state.index, &self.lang_def).or_else(|| if cursor_state.index > 0 { self.my_editor.find_matching_bracket(cursor_state.line, cursor_state.index - 1, &self.lang_def) } else { None });
        let mut total_h = 0.0; self.editor.with_buffer(|b| { for r in b.layout_runs() { if !self.is_line_hidden(r.line_i) { total_h += r.line_height; } } });
        self.scroll_y = self.scroll_y.clamp(0.0, (total_h - (rect.height() - FOOTER_HEIGHT * SCALE) + 100.0).max(0.0));
        
        let theme = &self.theme;
        let mut cp = Paint::default(); cp.set_color_rgba8(theme.current_line.r(), theme.current_line.g(), theme.current_line.b(), theme.current_line.a());
        let mut mp = Paint::default(); mp.set_color_rgba8(theme.match_highlight.r(), theme.match_highlight.g(), theme.match_highlight.b(), 100);
        let mut last_para = None;

        let my_editor = &self.my_editor;
        let _folds = &my_editor.folds;
        let collapsed = &my_editor.collapsed_starts;
        let digit_cache = &self.digit_cache;
        let show_whitespace = self.show_whitespace;
        let scroll_y = self.scroll_y;
        let mut glyph_cache = std::mem::take(&mut self.glyph_cache);

        let mut v_shift = 0.0;
        let mut initial_top = None;
        self.editor.with_buffer(|buffer| {
            for run in buffer.layout_runs() {
                if initial_top.is_none() { initial_top = Some(run.line_top); }
                let itop = initial_top.unwrap_or(0.0);
                
                let i = run.line_i;
                if self.is_line_hidden(i) {
                    v_shift += run.line_height;
                    continue;
                }
                let metrics = &self.metrics;
                let lh = run.line_height;
                let adj_lt = run.line_top - itop - v_shift;
                // Vertically center baseline in line height
                let centering_v = (lh - metrics.font_size) / 2.0;
                let adj_ly = adj_lt + metrics.font_size - centering_v + 4.0 * SCALE;
                
                let cyo = y_off - scroll_y;
                let v_t = adj_lt + cyo;
                
                if v_t + lh < rect.top() { continue; }
                if v_t > rect.bottom() { break; }

                let glyphs = run.glyphs;
                let text = run.text;
                
                // Aura Glow: Current line highlight with subtle gradient/border
                if i == cursor_state.line { 
                    let mut cp = Paint::default(); 
                    cp.set_color_rgba8(theme.current_line.r(), theme.current_line.g(), theme.current_line.b(), 100);
                    pixmap.fill_rect(Rect::from_xywh(rect.left() + GUTTER_WIDTH * SCALE, cyo + adj_lt, rect.width() - (GUTTER_WIDTH + MINIMAP_WIDTH) * SCALE, lh).unwrap(), &cp, Transform::identity(), None); 
                    // Subtle bottom border for the glow effect
                    let mut gp = Paint::default(); gp.set_color_rgba8(theme.kw.r(), theme.kw.g(), theme.kw.b(), 40);
                    pixmap.fill_rect(Rect::from_xywh(rect.left() + GUTTER_WIDTH * SCALE, cyo + adj_lt + lh - 1.0, rect.width() - (GUTTER_WIDTH + MINIMAP_WIDTH) * SCALE, 1.0).unwrap(), &gp, Transform::identity(), None);
                }
                if let Some(st) = &selected_text { 
                    if st.len() > 1 && !st.contains('\n') { 
                        let mut start = 0; 
                        while let Some(pos) = text[start..].find(st) { 
                            let real_pos = start + pos; let mut g_x = 0.0; let mut g_w = 0.0; 
                            for g in glyphs { if g.start >= real_pos && g.start < real_pos + st.len() { if g_w == 0.0 { g_x = g.x; } g_w += g.w; } } 
                            if g_w > 0.0 { pixmap.fill_rect(Rect::from_xywh(x_off + g_x, cyo + adj_lt, g_w, lh).unwrap(), &mp, Transform::identity(), None); } start = real_pos + 1; 
                        } 
                    } 
                }
                
                for diag in &my_editor.diagnostics {
                    if diag.range.start.line as usize == i {
                        let mut g_x = 0.0; let mut g_w = 0.0;
                        for g in glyphs { if g.start >= diag.range.start.character as usize && (diag.range.start.line != diag.range.end.line || g.start < diag.range.end.character as usize) { if g_w == 0.0 { g_x = g.x; } g_w += g.w; } }
                        if g_w > 0.0 { 
                            let mut pb = PathBuilder::new();
                            let mut cx = x_off + g_x; let step = 3.0 * SCALE; let amp = 1.5 * SCALE; let sy = cyo + adj_lt + lh - 2.0 * SCALE;
                            pb.move_to(cx, sy);
                            while cx < x_off + g_x + g_w { pb.quad_to(cx + step/2.0, sy + amp, cx + step, sy); pb.quad_to(cx + step * 1.5, sy - amp, cx + step * 2.0, sy); cx += step * 2.0; }
                            if let Some(path) = pb.finish() {
                                let mut paint = Paint::default(); paint.set_color_rgba8(theme.diagnostic_error.r(), theme.diagnostic_error.g(), theme.diagnostic_error.b(), 200);
                                pixmap.stroke_path(&path, &paint, &Stroke { width: 1.0 * SCALE, ..Default::default() }, Transform::identity(), None);
                            }
                        }
                    }
                }

                let mut gp_paint = Paint::default(); gp_paint.set_color_rgba8(theme.guide.r(), theme.guide.g(), theme.guide.b(), theme.guide.a());
                let tw = 4.0 * 8.4 * (metrics.font_size / 14.0) * SCALE; let ls = text.chars().take_while(|c| c.is_whitespace()).count();
                for j in 1..=(ls/4) { pixmap.fill_rect(Rect::from_xywh(x_off + (j as f32 * tw), cyo + adj_lt, 1.0, lh).unwrap(), &gp_paint, Transform::identity(), None); }
                if last_para != Some(i) {
                    let s = format!("{}", i + 1); let mut dx = (rect.left() + GUTTER_WIDTH * SCALE) as i32 - 15;
                    for ch in s.chars().rev() { if let Some(d) = ch.to_digit(10) { if (d as usize) < digit_cache.len() { let cg = &digit_cache[d as usize]; pixmap.draw_pixmap(dx - cg.pixmap.width() as i32 + cg.left, (cyo + adj_ly) as i32 - cg.top, cg.pixmap.as_ref(), &PixmapPaint::default(), Transform::identity(), None); dx -= 10 * SCALE as i32; } } }
                    if my_editor.folds.iter().any(|(s, _)| *s == i) { 
                        let col_folded = collapsed.contains(&i); 
                        App::draw_ui_text(pixmap, fs, sc, if col_folded { "+" } else { "-" }, rect.left() + 5.0 * SCALE, cyo + adj_ly - (12.0 * SCALE), if col_folded { theme.kw } else { Color::rgb(0x85, 0x85, 0x85) }); 
                    }
                    last_para = Some(i);
                }
                if let Some((ss, se)) = selection { 
                    if let Some((hx, hw)) = run.highlight(ss, se) { 
                        let mut sp_paint = Paint::default(); sp_paint.set_color_rgba8(theme.selection.r(), theme.selection.g(), theme.selection.b(), theme.selection.a()); 
                        pixmap.fill_rect(Rect::from_xywh(x_off + hx, cyo + adj_lt, hw, lh).unwrap(), &sp_paint, Transform::identity(), None); 
                    } 
                }
                for g in glyphs {
                    let ip = partner.map(|(pl, pi)| i == pl && g.start == pi).unwrap_or(false);
                    let pg = g.physical((x_off, cyo + adj_ly), 1.0); let gc = g.color_opt.unwrap_or(theme.text);
                    if show_whitespace { if text[g.start..g.end].starts_with(' ') { let mut wp = Paint::default(); wp.set_color_rgba8(80, 80, 80, 255); pixmap.fill_rect(Rect::from_xywh(x_off + g.x + g.w/2.0 - 1.0, cyo + adj_ly - (lh*0.25), 2.0, 2.0).unwrap(), &wp, Transform::identity(), None); } else if text[g.start..g.end].starts_with('\t') { let mut wp = Paint::default(); wp.set_color_rgba8(80, 80, 80, 255); pixmap.fill_rect(Rect::from_xywh(x_off + g.x + 2.0, cyo + adj_ly - (lh*0.25), g.w - 4.0, 1.0).unwrap(), &wp, Transform::identity(), None); } }
                    let cg = glyph_cache.entry((pg.cache_key, gc, ip)).or_insert_with(|| {
                        if let Some(im) = sc.get_image(fs, pg.cache_key) {
                            let mut p = Pixmap::new(im.placement.width.max(1), im.placement.height.max(1)).unwrap(); let (r, g, b, a) = if ip { (theme.bracket.r(), theme.bracket.g(), theme.bracket.b(), theme.bracket.a()) } else { (gc.r(), gc.g(), gc.b(), gc.a()) };
                            for (idx, &al) in im.data.iter().enumerate() { let af = (al as f32 / 255.0) * (a as f32 / 255.0); p.pixels_mut()[idx] = ColorU8::from_rgba((r as f32 * af) as u8, (g as f32 * af) as u8, (b as f32 * af) as u8, (255.0 * af) as u8).premultiply(); }
                            CachedGlyph { pixmap: p, left: im.placement.left, top: im.placement.top }
                        } else { CachedGlyph { pixmap: Pixmap::new(1, 1).unwrap(), left: 0, top: 0 } }
                    });
                    pixmap.draw_pixmap(pg.x + cg.left, pg.y - cg.top, cg.pixmap.as_ref(), &PixmapPaint::default(), Transform::identity(), None);
                    if ip { let mut bp = Paint::default(); bp.set_color_rgba8(theme.bracket.r(), theme.bracket.g(), theme.bracket.b(), theme.bracket.a()); pixmap.fill_rect(Rect::from_xywh(x_off + g.x, cyo + adj_lt + lh - 2.0, g.w, 2.0).unwrap(), &bp, Transform::identity(), None); }
                }

                // Match Highlighting: highlight occurrences of word under cursor
                if !selected_text.is_some() {
                    let cli = cursor_state.line; let cur = cursor_state.index;
                    let line_text = buffer.lines[cli].text();
                    
                    // Safely find word boundaries using byte offsets
                    let mut start = cur;
                    while start > 0 {
                        if let Some(c) = line_text[..start].chars().next_back() {
                            if c.is_alphanumeric() || c == '_' {
                                start -= c.len_utf8();
                                continue;
                            }
                        }
                        break;
                    }
                    let mut end = cur;
                    while end < line_text.len() {
                        if let Some(c) = line_text[end..].chars().next() {
                            if c.is_alphanumeric() || c == '_' {
                                end += c.len_utf8();
                                continue;
                            }
                        }
                        break;
                    }

                    if end > start {
                        let word = &line_text[start..end];
                        if word.len() > 1 && !self.lang_def.keywords.contains(word) {
                            let mut s = 0;
                            while s < line_text.len() {
                                if let Some(pos) = line_text[s..].find(word) {
                                    let rp = s + pos; let mut gx = 0.0; let mut gw = 0.0;
                                    for g in glyphs { if g.start >= rp && g.start < rp + word.len() { if gw == 0.0 { gx = g.x; } gw += g.w; } }
                                    if gw > 0.0 { 
                                        let mut mp = Paint::default(); mp.set_color_rgba8(theme.match_highlight.r(), theme.match_highlight.g(), theme.match_highlight.b(), 80);
                                        pixmap.fill_rect(Rect::from_xywh(x_off + gx, cyo + adj_lt, gw, lh).unwrap(), &mp, Transform::identity(), None); 
                                    }
                                    s = rp + word.len();
                                } else {
                                    break;
                                }
                            }
                        }
                    }
                }
            }
        });
        self.glyph_cache = glyph_cache;
        let minimap_rect = Rect::from_xywh(rect.right() - MINIMAP_WIDTH * SCALE, rect.top(), MINIMAP_WIDTH * SCALE, rect.height() - FOOTER_HEIGHT * SCALE).unwrap();
        let mx = minimap_rect.left();
        
        let m_w = minimap_rect.width() as u32;
        let m_h = minimap_rect.height() as u32;

        if self.minimap_needs_redraw || self.minimap_pixmap.as_ref().map(|p| p.width() != m_w || p.height() != m_h).unwrap_or(true) {
            let mut mp = Pixmap::new(m_w.max(1), m_h.max(1)).unwrap();
            let to_skia = |c: Color| SkiaColor::from_rgba8(c.r(), c.g(), c.b(), c.a());
            mp.fill(to_skia(theme.minimap_bg));
            
            let m_step = (minimap_rect.height() / self.my_editor.line_tokens.len().max(1) as f32).min(2.5 * SCALE); 
            let mut m_y = 0.0;
            for li in 0..self.my_editor.line_tokens.len() {
                if self.is_line_hidden(li) { continue; }
                if let Some(tks) = self.my_editor.line_tokens.get(li) {
                    let mut xp = 2.0;
                    for t in tks {
                        let tc = match t.kind {
                            TokenKind::Identifier => self.theme.kw,
                            TokenKind::String => self.theme.string,
                            TokenKind::LineComment | TokenKind::BlockComment => self.theme.comment,
                            TokenKind::Number => self.theme.number,
                            _ => self.theme.guide
                        };
                        let mut tp = Paint::default();
                        tp.set_color_rgba8(tc.r(), tc.g(), tc.b(), 0xaa);
                        let w = ((t.end - t.start) as f32 * 0.4 * SCALE).min(minimap_rect.width() - xp);
                        if w > 0.0 {
                            mp.fill_rect(Rect::from_xywh(xp, m_y, w, (m_step * 0.7).max(1.0)).unwrap(), &tp, Transform::identity(), None);
                        }
                        xp += w + 1.0;
                        if xp >= minimap_rect.width() { break; }
                    }
                }
                m_y += m_step;
                if m_y > minimap_rect.height() { break; }
            }
            self.minimap_pixmap = Some(mp);
            self.minimap_needs_redraw = false;
        }

        if let Some(mp) = &self.minimap_pixmap {
            pixmap.draw_pixmap(mx as i32, rect.top() as i32, mp.as_ref(), &PixmapPaint::default(), Transform::identity(), None);
        }

        let v_h = ((rect.height() - FOOTER_HEIGHT * SCALE) / total_h.max(1.0)) * (rect.height() - FOOTER_HEIGHT * SCALE);
        let v_y = (self.scroll_y / total_h.max(1.0)) * (rect.height() - FOOTER_HEIGHT * SCALE);
        let mut vp = Paint::default(); vp.set_color_rgba8(255, 255, 255, 17);
        pixmap.fill_rect(Rect::from_xywh(mx, rect.top() + v_y.min(rect.height() - FOOTER_HEIGHT * SCALE - v_h), MINIMAP_WIDTH * SCALE, v_h).unwrap(), &vp, Transform::identity(), None);
        if self.is_search_open {
             let mut sep = Paint::default(); sep.set_color_rgba8(51, 51, 51, 250);
             let eh = if self.is_replace_open { 60.0 } else { 30.0 } * SCALE;
             pixmap.fill_rect(Rect::from_xywh(rect.right() - 260.0 * SCALE, rect.top() + 10.0 * SCALE, 160.0 * SCALE, eh).unwrap(), &sep, Transform::identity(), None);
             App::draw_ui_text(pixmap, fs, sc, &format!("Find: {}", self.search_query), rect.right() - 250.0 * SCALE, rect.top() + 14.0 * SCALE, self.theme.text);
             if self.is_replace_open { App::draw_ui_text(pixmap, fs, sc, &format!("Replace: {}", self.replace_query), rect.right() - 250.0 * SCALE, rect.top() + 44.0 * SCALE, Color::rgb(206, 145, 120)); }
        }
        if let Some((pos, items)) = &self.context_menu {
            let mut mep = Paint::default(); mep.set_color_rgba8(45, 45, 45, 255);
            pixmap.fill_rect(Rect::from_xywh(pos.0, pos.1, 100.0 * SCALE, (items.len() as f32 * 25.0) * SCALE).unwrap(), &mep, Transform::identity(), None);
            for (i, item) in items.iter().enumerate() { App::draw_ui_text(pixmap, fs, sc, item, pos.0 + 10.0 * SCALE, pos.1 + (i as f32 * 25.0 + 5.0) * SCALE, self.theme.text); }
        }
        if let Some((cx, cy)) = self.editor.cursor_position() {
             let mut cy_final = None;
             self.editor.with_buffer(|b| {
                 let mut v_sh = 0.0;
                 let mut i_top = None;
                 for r in b.layout_runs() {
                     if i_top.is_none() { i_top = Some(r.line_top); }
                     if self.is_line_hidden(r.line_i) { v_sh += r.line_height; continue; }
                     if cy >= r.line_top as i32 && cy < (r.line_top + r.line_height) as i32 {
                         let it = i_top.unwrap_or(0.0);
                         let l_h = r.line_height;
                         let c_v = (l_h - self.metrics.font_size) / 2.0;
                         // Center characters in rectangle
                         cy_final = Some(y_off - self.scroll_y + (r.line_top - it - v_sh) + c_v);
                         break;
                     }
                 }
             });
             if let Some(final_y) = cy_final {
                 let mut cp = Paint::default(); cp.set_color_rgba8(255, 255, 255, 255);
                 pixmap.fill_rect(Rect::from_xywh(x_off + cx as f32, final_y, 2.0, self.metrics.line_height).unwrap(), &cp, Transform::identity(), None);
             }
        }
        
        // 6. Context Menu Overlay
        if let Some(((mx, my), items)) = &self.context_menu {
            let mut bg = Paint::default(); bg.set_color_rgba8(40, 40, 45, 255);
            let mw = 120.0 * SCALE; let mh = items.len() as f32 * 25.0 * SCALE;
            pixmap.fill_rect(Rect::from_xywh(*mx * SCALE, *my * SCALE, mw, mh).unwrap(), &bg, Transform::identity(), None);
            for (i, item) in items.iter().enumerate() {
                App::draw_ui_text(pixmap, fs, sc, item, *mx * SCALE + 10.0 * SCALE, *my * SCALE + (i as f32 * 25.0 + 5.0) * SCALE, Color::rgb(220, 220, 220));
            }
        }
    }

    pub fn handle_mouse(&mut self, fs: &mut FontSystem, x: f32, y: f32, rect: Rect, click: Option<(u32, MouseButton, winit::event::Modifiers)>, clipboard: &mut Option<Clipboard>) {
        if x < rect.left() || x > rect.right() || y < rect.top() || y > rect.bottom() { return; }
        if let Some((_, MouseButton::Right, _)) = click { self.context_menu = Some(((x, y), vec!["Copy".to_string(), "Paste".to_string(), "Cut".to_string(), "Select All".to_string(), "Find".to_string(), "Replace".to_string()])); return; }
        if let Some((pos, _)) = self.context_menu {
            let menu_h = (6.0 * 25.0) * SCALE;
            if x >= pos.0 && x <= pos.0 + 100.0 * SCALE && y >= pos.1 && y <= pos.1 + menu_h {
                let idx = ((y - pos.1) / (25.0 * SCALE)) as usize; self.context_menu = None;
                match idx { 
                    0 => { if let Some(t) = self.editor.copy_selection() { if let Some(cb) = clipboard { let _ = cb.set_text(t); } } },
                    1 => { if let Some(cb) = clipboard { if let Ok(t) = cb.get_text() { for ch in t.chars() { self.editor.action(fs, Action::Insert(ch)); } self.needs_reshape = true; self.sync(); } } },
                    2 => { if let Some(t) = self.editor.copy_selection() { if let Some(cb) = clipboard { let _ = cb.set_text(t); } self.editor.action(fs, Action::Delete); self.needs_reshape = true; self.sync(); } },
                    3 => { self.editor.action(fs, Action::Motion(cosmic_text::Motion::BufferStart)); let mut ly = 0.0; self.editor.with_buffer(|b| if let Some(r) = b.layout_runs().last() { ly = r.line_top + r.line_height; }); self.editor.action(fs, Action::Drag { x: 999999, y: ly as i32 }); },
                    4 => { self.is_search_open = true; self.search_query.clear(); },
                    5 => { self.is_search_open = true; self.is_replace_open = true; self.search_query.clear(); self.replace_query.clear(); },
                    _ => {} 
                } return;
            } self.context_menu = None;
        }
        if y >= rect.top() && y < rect.bottom() && x > rect.right() - MINIMAP_WIDTH * SCALE {
             let mry = (y - rect.top()) / rect.height(); let mut th = 0.0;
             self.editor.with_buffer(|b| { for r in b.layout_runs() { if !self.is_line_hidden(r.line_i) { th += r.line_height; } } });
             self.scroll_y = (mry * th).max(0.0); return;
        }
        let (x_off, y_off) = self.get_offsets(rect);

        if let Some((1, MouseButton::Left, _)) = click { 
            if x < x_off && x > rect.left() { 
                let mut fl = None; 
                self.editor.with_buffer(|b| { 
                    let mut i_top = None;
                    let mut v_sh = 0.0;
                    for r in b.layout_runs() { 
                        if i_top.is_none() { i_top = Some(r.line_top); }
                        if self.is_line_hidden(r.line_i) { v_sh += r.line_height; continue; }
                        let vy = y_off - self.scroll_y + (r.line_top - i_top.unwrap_or(0.0) - v_sh); 
                        if y >= vy && y < vy + r.line_height { fl = Some(r.line_i); break; } 
                    } 
                }); 
                if let Some(li) = fl { self.my_editor.toggle_fold(li); return; } 
            } 
        }
        let mut final_ts = 0.0;
        self.editor.with_buffer(|b| { 
            let mut i_top = None;
            let mut v_sh = 0.0;
            for r in b.layout_runs() { 
                if i_top.is_none() { i_top = Some(r.line_top); }
                let it = i_top.unwrap_or(0.0);
                if self.is_line_hidden(r.line_i) { v_sh += r.line_height; continue; } 
                let vy = y_off - self.scroll_y + (r.line_top - it - v_sh);
                if y >= vy && y < vy + r.line_height { 
                    final_ts = it + v_sh; break; 
                } 
            } 
        });
        let ex = (x - x_off) as i32; let ey = (y - y_off + final_ts + self.scroll_y) as i32;
        if let Some((count, _, mods)) = click {
            if !mods.state().shift_key() && count == 1 { self.editor.action(fs, Action::Escape); }
            match count { 1 => self.editor.action(fs, Action::Click { x: ex, y: ey }), 2 => self.editor.action(fs, Action::DoubleClick { x: ex, y: ey }), 3 => self.editor.action(fs, Action::TripleClick { x: ex, y: ey }), _ => {} }
        } else { self.editor.action(fs, Action::Drag { x: ex, y: ey }); }
    }

    pub fn sync(&mut self) { 
        if self.needs_reshape { 
            let mut t = String::new();
            self.editor.with_buffer(|b| {
                for (i, line) in b.lines.iter().enumerate() {
                    if i > 0 { t.push('\n'); }
                    t.push_str(line.text());
                }
            });
            self.my_editor.rope = ropey::Rope::from_str(&t); 
            self.my_editor.retokenize_all(&self.lang_def); 
            self.reapply_highlighting();
            self.minimap_needs_redraw = true;
        } 
    }
}

struct App {
    window: Option<Arc<Window>>,
    context: Option<Context<Arc<Window>>>,
    surface: Option<Surface<Arc<Window>, Arc<Window>>>,
    font_system: FontSystem,
    swash_cache: SwashCache,
    pixmap: Option<Pixmap>,
    tabs: Vec<Tab>,
    active_tab: usize,
    tree_view: TreeView,
    all_languages: Vec<String>,
    current_lang: String,
    lang_dropdown: Option<Dropdown>,
    theme_dropdown: Option<Dropdown>,
    clipboard: Option<Clipboard>,
    modifiers: winit::event::Modifiers,
    last_click_time: Instant,
    click_count: u32,
    mouse_pos: (f32, f32),
    explorer_width: f32,
    is_dragging_splitter: bool,
    hovering_splitter: bool,
    hovering_tab_close: Option<usize>,
    is_dragging: bool,
    needs_redraw: bool,
    last_lsp_update: Instant,
    pending_lsp_update: bool,
    lsp: Arc<LspClient>,
    is_quick_open: bool,
    quick_open_query: String,
    tab_scroll_x: f32,
    current_theme_idx: usize,
    breadcrumb_rects: Vec<(Rect, String)>,
    keybindings: Vec<Keybinding>,
}

impl App {
    fn new(_my_editor: MyEditor) -> Self {
        let mut langs = Vec::new(); 
        // Try multiple paths to find the basic-languages folder
        let search_paths = ["crates/code_editor/basic-languages", "basic-languages", "../code_editor/basic-languages"];
        for path in &search_paths {
            if let Ok(es) = std::fs::read_dir(path) { 
                for e in es.flatten() { 
                    if e.file_type().map(|t| t.is_dir()).unwrap_or(false) { 
                        if let Some(n) = e.file_name().to_str() { langs.push(n.to_string()); } 
                    } 
                } 
                if !langs.is_empty() { break; }
            }
        }
        
        // Final fallback if no languages found
        if langs.is_empty() {
            println!("WARNING: Could not find basic-languages folder, using hardcoded fallback.");
            langs = vec!["rust".into(), "javascript".into(), "typescript".into(), "python".into(), "text".into()];
        }
        langs.sort();
        let root_dir = std::env::current_dir().unwrap_or_else(|_| std::path::PathBuf::from("/"));
        let _root_uri = format!("file://{}", root_dir.to_string_lossy());
        Self { 
            window: None, context: None, surface: None, font_system: FontSystem::new(), swash_cache: SwashCache::new(), pixmap: None, tabs: Vec::new(), active_tab: 0, 
            tree_view: TreeView::new(".", 2.0), all_languages: langs, current_lang: "rust".to_string(), lang_dropdown: None, theme_dropdown: None, clipboard: Clipboard::new().ok(), 
            modifiers: winit::event::Modifiers::default(), last_click_time: Instant::now(), click_count: 0, mouse_pos: (0.0, 0.0), 
            explorer_width: EXPLORER_WIDTH, is_dragging_splitter: false, hovering_splitter: false, hovering_tab_close: None,
            is_dragging: false, needs_redraw: true,
            last_lsp_update: Instant::now(), pending_lsp_update: false,
            lsp: Arc::new(LspClient::new()),
            is_quick_open: false,
            quick_open_query: String::new(),
            tab_scroll_x: 0.0,
            current_theme_idx: 0,
            breadcrumb_rects: Vec::new(),
            keybindings: {
                let paths = ["keybindings.json", "crates/code_editor/keybindings.json"];
                let mut kb = Vec::new();
                for p in paths {
                    if let Ok(s) = fs::read_to_string(p) {
                        if let Ok(parsed) = serde_json::from_str::<Vec<Keybinding>>(&s) {
                            kb = parsed;
                            break;
                        }
                    }
                }
                if kb.is_empty() {
                    kb = vec![
                        Keybinding { key: "z".into(), cmd: true, shift: false, alt: false, action: "Undo".into() },
                        Keybinding { key: "z".into(), cmd: true, shift: true, alt: false, action: "Redo".into() },
                        Keybinding { key: "a".into(), cmd: true, shift: false, alt: false, action: "SelectAll".into() },
                        Keybinding { key: "s".into(), cmd: true, shift: false, alt: false, action: "Save".into() },
                        Keybinding { key: "f".into(), cmd: true, shift: false, alt: false, action: "Find".into() },
                        Keybinding { key: "h".into(), cmd: true, shift: false, alt: false, action: "Replace".into() },
                        Keybinding { key: "/".into(), cmd: true, shift: false, alt: false, action: "ToggleComment".into() },
                        Keybinding { key: "Tab".into(), cmd: false, shift: false, alt: false, action: "Indent".into() },
                        Keybinding { key: "Tab".into(), cmd: false, shift: true, alt: false, action: "Unindent".into() },
                        Keybinding { key: "ArrowUp".into(), cmd: true, shift: false, alt: false, action: "MoveBufferStart".into() },
                        Keybinding { key: "ArrowDown".into(), cmd: true, shift: false, alt: false, action: "MoveBufferEnd".into() },
                        Keybinding { key: "ArrowLeft".into(), cmd: true, shift: false, alt: false, action: "MoveLineStart".into() },
                        Keybinding { key: "ArrowRight".into(), cmd: true, shift: false, alt: false, action: "MoveLineEnd".into() },
                        Keybinding { key: "ArrowLeft".into(), cmd: false, shift: false, alt: true, action: "MoveWordLeft".into() },
                        Keybinding { key: "ArrowRight".into(), cmd: false, shift: false, alt: true, action: "MoveWordRight".into() },
                    ];
                }
                kb
            },
        }
    }
    pub fn active_theme(&self) -> Theme {
        match self.current_theme_idx {
            0 => Theme::silicon_green(),
            1 => Theme::cloud_blue(),
            2 => Theme::coffee_cream(),
            3 => Theme::sakura_pink(),
            4 => Theme::one_dark(),
            5 => Theme::monokai(),
            6 => Theme::frost_light(),
            7 => Theme::solarized_light(),
            8 => Theme::midnight(),
            9 => Theme::aura(),
            10 => Theme::veridian(),
            11 => Theme::rose(),
            12 => Theme::cyber(),
            13 => Theme::titanium(),
            14 => Theme::indigo_night(),
            _ => Theme::one_dark(),
        }
    }

    pub fn get_theme_name(&self) -> &str {
        match self.current_theme_idx { 
            0 => "Silicon Green", 1 => "Cloud Blue", 2 => "Coffee Cream", 3 => "Sakura Pink", 
            4 => "One Dark", 5 => "Monokai", 6 => "Frost Light", 7 => "Solarized Light", 
            8 => "Midnight", 9 => "Aura", 10 => "Veridian", 11 => "Rose",
            12 => "Cyber", 13 => "Titanium", 14 => "Indigo Night", _ => "One Dark" 
        }
    }

    fn render(&mut self) {
        // Debounce LSP Update
        if self.pending_lsp_update && self.last_lsp_update.elapsed().as_millis() > 300 {
            if let Some(tab) = self.tabs.get_mut(self.active_tab) {
                let text = tab.widget.my_editor.rope.to_string();
                let uri = tab.path.clone().unwrap_or_else(|| format!("file:///Users/youness/www/html/vybe/{}", tab.name));
                tab.widget.lsp.send(LspRequest::Change(text, uri));
            }
            self.pending_lsp_update = false;
        }

        let theme = self.active_theme();
        let theme_name = self.get_theme_name().to_string();
        let (surf, pix) = match (&mut self.surface, &mut self.pixmap) { (Some(s), Some(p)) => (s, p), _ => return };
        
        let to_skia = |c: Color| SkiaColor::from_rgba8(c.r(), c.g(), c.b(), c.a());
        pix.fill(to_skia(theme.bg));

        // 1. Sidebar (Project Explorer)
        let mut sp = Paint::default(); sp.set_color_rgba8(theme.sidebar_bg.r(), theme.sidebar_bg.g(), theme.sidebar_bg.b(), theme.sidebar_bg.a());
        pix.fill_rect(Rect::from_xywh(0.0, 0.0, self.explorer_width * SCALE, pix.height() as f32).unwrap(), &sp, Transform::identity(), None);
        
        // Vertical marker for Sidebar header
        App::draw_ui_text(pix, &mut self.font_system, &mut self.swash_cache, "PROJECT EXPLORER", 10.0 * SCALE, 10.0 * SCALE, theme.sidebar_text);
        
        // Sync Sidebar Selection
        if let Some(tab) = self.tabs.get(self.active_tab) {
             if let Some(path) = &tab.path {
                  self.tree_view.reveal_path(path);
             }
        }
        self.tree_view.render(pix, &mut self.font_system, &mut self.swash_cache, 0.0, TAB_BAR_HEIGHT * SCALE, self.explorer_width * SCALE, theme.sidebar_text, (theme.selection.r(), theme.selection.g(), theme.selection.b(), theme.selection.a()));

        // 1b. Splitter
        let mut slp = Paint::default();
        if self.is_dragging_splitter { slp.set_color_rgba8(0,122,204,255); }
        else if self.hovering_splitter { slp.set_color_rgba8(theme.splitter_bg.r(), theme.splitter_bg.g(), theme.splitter_bg.b(), 255); }
        else { slp.set_color_rgba8(theme.splitter_bg.r(), theme.splitter_bg.g(), theme.splitter_bg.b(), 255); }
        pix.fill_rect(Rect::from_xywh(self.explorer_width * SCALE, 0.0, SPLITTER_WIDTH * SCALE, pix.height() as f32).unwrap(), &slp, Transform::identity(), None);
        
        // Splitter separation line
        let mut lp = Paint::default(); lp.set_color_rgba8(theme.splitter_bg.r().saturating_add(20), theme.splitter_bg.g().saturating_add(20), theme.splitter_bg.b().saturating_add(20), 255);
        pix.fill_rect(Rect::from_xywh((self.explorer_width + SPLITTER_WIDTH) * SCALE, 0.0, 1.0 * SCALE, pix.height() as f32).unwrap(), &lp, Transform::identity(), None);

        // 2. Tab Bar
        let ed_start_x = (self.explorer_width + SPLITTER_WIDTH + 1.0) * SCALE;
        let mut tp = Paint::default(); tp.set_color_rgba8(theme.tab_bar_bg.r(), theme.tab_bar_bg.g(), theme.tab_bar_bg.b(), theme.tab_bar_bg.a());
        pix.fill_rect(Rect::from_xywh(ed_start_x, 0.0, pix.width() as f32 - ed_start_x, TAB_BAR_HEIGHT * SCALE).unwrap(), &tp, Transform::identity(), None);

            let mut tx_off = ed_start_x + self.tab_scroll_x;
            for i in 0..self.tabs.len() {
                if tx_off + 160.0 * SCALE < ed_start_x { tx_off += 160.0 * SCALE; continue; }
                if tx_off > pix.width() as f32 { break; }
                
                let active = i == self.active_tab;
                let tw = 160.0 * SCALE;
                
                // Render Background & Underline
                if active {
                    let mut ap = Paint::default(); ap.set_color_rgba8(theme.active_tab_bg.r(), theme.active_tab_bg.g(), theme.active_tab_bg.b(), theme.active_tab_bg.a());
                    pix.fill_rect(Rect::from_xywh(tx_off, 0.0, tw, TAB_BAR_HEIGHT * SCALE).unwrap(), &ap, Transform::identity(), None);
                    
                    let mut up = Paint::default(); up.set_color_rgba8(theme.kw.r(), theme.kw.g(), theme.kw.b(), 255);
                    pix.fill_rect(Rect::from_xywh(tx_off, (TAB_BAR_HEIGHT - 2.0) * SCALE, tw, 2.0 * SCALE).unwrap(), &up, Transform::identity(), None);
                } else {
                    let mut ip = Paint::default(); ip.set_color_rgba8(theme.inactive_tab_bg.r(), theme.inactive_tab_bg.g(), theme.inactive_tab_bg.b(), theme.inactive_tab_bg.a());
                    pix.fill_rect(Rect::from_xywh(tx_off, 0.0, tw, TAB_BAR_HEIGHT * SCALE).unwrap(), &ip, Transform::identity(), None);
                }

                // Get tab properties for name calculation
                let (is_sticky, name, is_modified) = {
                    let t = &self.tabs[i];
                    (t.is_sticky, t.name.clone(), t.is_modified)
                };
                let name_str = if is_sticky { name } else { format!("{} [P]", name) };
                let col = if active { theme.active_tab_text } else { theme.inactive_tab_text };

                // Tab Text Caching & Rendering
                let tab_mut = &mut self.tabs[i];
                if tab_mut.buffer.is_none() {
                    let mut lab = Buffer::new(&mut self.font_system, Metrics::new(14.0,20.0).scale(SCALE));
                    lab.set_text(&mut self.font_system, &name_str, &Attrs::new().family(Family::Monospace).color(col), Shaping::Advanced, None);
                    lab.shape_until_scroll(&mut self.font_system, false);
                    tab_mut.buffer = Some(lab);
                }
                if let Some(lab) = &tab_mut.buffer {
                    for r in lab.layout_runs() {
                        for g in r.glyphs {
                            let pg = g.physical((tx_off + 10.0 * SCALE, 10.0 * SCALE + r.line_y), 1.0);
                            if let Some(im) = self.swash_cache.get_image(&mut self.font_system, pg.cache_key) {
                                let mut p = Pixmap::new(im.placement.width.max(1), im.placement.height.max(1)).unwrap();
                                let (cr, cg, cb, ca) = (col.r(), col.g(), col.b(), col.a());
                                for (idx, &al) in im.data.iter().enumerate() {
                                    let af = (al as f32 / 255.0) * (ca as f32 / 255.0);
                                    p.pixels_mut()[idx] = ColorU8::from_rgba((cr as f32 * af) as u8, (cg as f32 * af) as u8, (cb as f32 * af) as u8, (255.0 * af) as u8).premultiply();
                                }
                                pix.draw_pixmap(pg.x + im.placement.left, pg.y - im.placement.top, p.as_ref(), &PixmapPaint::default(), Transform::identity(), None);
                            }
                        }
                    }
                }
                
                // Tab close button [X] or Modified dot [•]
                if is_modified {
                    App::draw_ui_text(pix, &mut self.font_system, &mut self.swash_cache, "•", tx_off + tw - 24.0 * SCALE, 10.0 * SCALE, Color::rgb(180, 180, 180));
                } else {
                    let is_close_hover = self.hovering_tab_close == Some(i);
                    App::draw_ui_text(pix, &mut self.font_system, &mut self.swash_cache, "×", tx_off + tw - 24.0 * SCALE, 10.0 * SCALE, if is_close_hover { Color::rgb(255, 100, 100) } else { Color::rgb(120,120,120) });
                }

                tx_off += tw;
            }

        // 2b. Breadcrumbs Removed (Unified in Status Bar)

        // 3. Active Editor
        if self.active_tab < self.tabs.len() {
             let ed_top = (TAB_BAR_HEIGHT + UI_BAR_HEIGHT) * SCALE;
             let rect = Rect::from_xywh(ed_start_x, ed_top, pix.width() as f32 - ed_start_x, pix.height() as f32 - (ed_top + FOOTER_HEIGHT * SCALE)).unwrap();
             
             while let Ok(evt) = self.lsp.rx.try_recv() {
                match evt {
                    LspEvent::Diagnostics(uri, diags) => { 
                        for t in &mut self.tabs {
                            let t_uri = t.path.clone().unwrap_or_else(|| format!("file:///Users/youness/www/html/vybe/{}", t.name));
                            if t_uri == uri {
                                t.widget.my_editor.diagnostics = diags;
                                self.needs_redraw = true;
                                break;
                            }
                        }
                    }
                    _ => {}
                }
             }
             let w = &mut self.tabs[self.active_tab].widget;
             
             // Update Editor Wrapping if changed
             w.editor.with_buffer_mut(|b| {
                 let wrap = if w.wrap_lines { cosmic_text::Wrap::Word } else { cosmic_text::Wrap::None };
                 if b.wrap() != wrap {
                     b.set_wrap(&mut self.font_system, wrap);
                     w.needs_reshape = true;
                 }
                 if w.wrap_lines {
                     b.set_size(&mut self.font_system, Some(rect.width() - (GUTTER_WIDTH + MINIMAP_WIDTH) * SCALE), Some(rect.height()));
                 } else {
                     b.set_size(&mut self.font_system, Some(999999.0), Some(999999.0));
                 }
             });

             w.render(pix, &mut self.font_system, &mut self.swash_cache, rect);
        }

        // 4. Footer (Enhanced)
        let mut fp = Paint::default(); fp.set_color_rgba8(theme.footer_bg.r(), theme.footer_bg.g(), theme.footer_bg.b(), theme.footer_bg.a());
        pix.fill_rect(Rect::from_xywh(0.0, pix.height() as f32 - FOOTER_HEIGHT * SCALE, pix.width() as f32, FOOTER_HEIGHT * SCALE).unwrap(), &fp, Transform::identity(), None);
        
        if let Some(tab) = self.tabs.get(self.active_tab) {
            let cursor = tab.widget.editor.cursor();
            let text = tab.widget.my_editor.rope.to_string();
            let line_endings = if text.contains("\r\n") { "CRLF" } else { "LF" };
            
            // segments for breadcrumbs and zoom indicator
            let path_str = tab.path.clone().unwrap_or_else(|| tab.name.clone());
            let segments: Vec<&str> = path_str.split(|c| c == '/' || c == '\\').filter(|s| !s.is_empty()).collect();
            
            let zoom_pct = (tab.widget.font_size / 14.0 * 100.0) as i32;
            let status_prefix = format!("Ln {}, Col {} | {}% | {} | UTF-8 | ", cursor.line + 1, cursor.index + 1, zoom_pct, line_endings);
            App::draw_ui_text(pix, &mut self.font_system, &mut self.swash_cache, &status_prefix, 10.0 * SCALE, pix.height() as f32 - FOOTER_HEIGHT * SCALE + 4.0 * SCALE, theme.footer_text);
            
            // Draw interactive breadcrumbs
            let mut current_x = 10.0 * SCALE + (status_prefix.len() as f32 * 8.4 * SCALE); // hardcoded approx char width
            self.breadcrumb_rects.clear();
            for (i, seg) in segments.iter().enumerate() {
                let seg_text = if i == segments.len() - 1 { seg.to_string() } else { format!("{} > ", seg) };
                let seg_width = seg_text.len() as f32 * 8.4 * SCALE;
                let rect = Rect::from_xywh(current_x, pix.height() as f32 - FOOTER_HEIGHT * SCALE, seg_width, FOOTER_HEIGHT * SCALE).unwrap();
                
                // Construct full path up to this segment
                let partial_path = segments[0..=i].join("/");
                self.breadcrumb_rects.push((rect, partial_path));
                
                App::draw_ui_text(pix, &mut self.font_system, &mut self.swash_cache, &seg_text, current_x, pix.height() as f32 - FOOTER_HEIGHT * SCALE + 4.0 * SCALE, theme.footer_text);
                current_x += seg_width;
            }

            let lang_label = format!("Language: {}", self.current_lang);
            let theme_label = format!("Theme: {}", theme_name);
            
            let label_x = pix.width() as f32 - (lang_label.len() as f32 * 9.0 + 20.0) * SCALE;
            let theme_x = label_x - (theme_label.len() as f32 * 9.0 + 30.0) * SCALE;

            App::draw_ui_text(pix, &mut self.font_system, &mut self.swash_cache, &lang_label, label_x, pix.height() as f32 - FOOTER_HEIGHT * SCALE + 4.0 * SCALE, theme.footer_text);
            App::draw_ui_text(pix, &mut self.font_system, &mut self.swash_cache, &theme_label, theme_x, pix.height() as f32 - FOOTER_HEIGHT * SCALE + 4.0 * SCALE, theme.footer_text);

            if let Some(dropdown) = &self.lang_dropdown {
                let (w, h) = dropdown.get_size();
                let menu_x = (pix.width() as f32 / SCALE - w - 20.0).max(10.0);
                let menu_y = (pix.height() as f32 / SCALE - FOOTER_HEIGHT - h - 10.0).max(10.0);
                dropdown.render(
                    pix, &mut self.font_system, &mut self.swash_cache, menu_x, menu_y,
                    (theme.sidebar_bg.r(), theme.sidebar_bg.g(), theme.sidebar_bg.b(), 255),
                    (theme.gutter_divider.r(), theme.gutter_divider.g(), theme.gutter_divider.b(), 255),
                    (theme.selection.r(), theme.selection.g(), theme.selection.b(), 100),
                    (theme.current_line.r(), theme.current_line.g(), theme.current_line.b(), 255),
                    theme.active_tab_text,
                    theme.inactive_tab_text
                );
            }
            if let Some(dropdown) = &self.theme_dropdown {
                let (w, h) = dropdown.get_size();
                let menu_x = (theme_x / SCALE - 10.0).max(10.0);
                // Edge Clamping
                let menu_x = menu_x.min(pix.width() as f32 / SCALE - w - 10.0).max(10.0);
                let menu_y = (pix.height() as f32 / SCALE - FOOTER_HEIGHT - h - 10.0).max(10.0);
                dropdown.render(
                    pix, &mut self.font_system, &mut self.swash_cache, menu_x, menu_y,
                    (theme.sidebar_bg.r(), theme.sidebar_bg.g(), theme.sidebar_bg.b(), 255),
                    (theme.gutter_divider.r(), theme.gutter_divider.g(), theme.gutter_divider.b(), 255),
                    (theme.selection.r(), theme.selection.g(), theme.selection.b(), 100),
                    (theme.current_line.r(), theme.current_line.g(), theme.current_line.b(), 255),
                    theme.active_tab_text,
                    theme.inactive_tab_text
                );
            }
        }

        // 5. Diagnostic Tooltip on Hover
        if let Some(tab) = self.tabs.get(self.active_tab) {
            let _w = &tab.widget;
            let _mx = self.mouse_pos.0; let _my = self.mouse_pos.1;
            // Diagnostic tooltips logic...
        }

        // Quick Open Overlay
        if self.is_quick_open {
            let mut o_p = Paint::default(); o_p.set_color_rgba8(30, 30, 35, 240);
            let o_w = 400.0 * SCALE;
            let o_h = 300.0 * SCALE;
            let o_x = (pix.width() as f32 - o_w) / 2.0;
            let o_y = 100.0 * SCALE;
            pix.fill_rect(Rect::from_xywh(o_x, o_y, o_w, o_h).unwrap(), &o_p, Transform::identity(), None);
            
            let mut b_p = Paint::default(); b_p.set_color_rgba8(80, 80, 90, 255);
            let mut pb = PathBuilder::new(); pb.push_rect(Rect::from_xywh(o_x, o_y, o_w, o_h).unwrap());
            if let Some(path) = pb.finish() { pix.stroke_path(&path, &b_p, &Stroke { width: 1.0 * SCALE, ..Default::default() }, Transform::identity(), None); }
            
            App::draw_ui_text(pix, &mut self.font_system, &mut self.swash_cache, &format!("Go to file: {}|", self.quick_open_query), o_x + 10.0 * SCALE, o_y + 10.0 * SCALE, Color::rgb(200, 200, 200));
            
            let matcher = SkimMatcherV2::default();
            let mut matches: Vec<(i64, usize, &String)> = self.tabs.iter().enumerate()
                .filter_map(|(idx, tab)| {
                    if self.quick_open_query.is_empty() {
                        Some((0, idx, &tab.name))
                    } else {
                        matcher.fuzzy_match(&tab.name, &self.quick_open_query).map(|score| (score, idx, &tab.name))
                    }
                })
                .collect();
            matches.sort_by_key(|m| -m.0); // Highest score first

            let mut i_y = o_y + 50.0 * SCALE;
            for (idx, (score, _tab_idx, name)) in matches.iter().take(10).enumerate() {
                let col = if idx == 0 { Color::rgb(0, 122, 204) } else { Color::rgb(200, 200, 200) };
                let display_text = if *score > 0 { format!("{} (score: {})", name, score) } else { name.to_string() };
                App::draw_ui_text(pix, &mut self.font_system, &mut self.swash_cache, &display_text, o_x + 20.0 * SCALE, i_y, col);
                i_y += 25.0 * SCALE;
            }
        }

        let mut buffer = surf.buffer_mut().unwrap();
        for (i, p) in pix.pixels().iter().enumerate() {
            buffer[i] = (p.red() as u32) << 16 | (p.green() as u32) << 8 | (p.blue() as u32);
        }
        buffer.present().unwrap();
        self.needs_redraw = false;
    }
    fn draw_ui_text(pix: &mut Pixmap, fs: &mut FontSystem, sc: &mut SwashCache, text: &str, x: f32, y: f32, col: Color) {
        let mut lab = Buffer::new(fs, Metrics::new(14.0,20.0).scale(SCALE)); lab.set_text(fs, text, &Attrs::new().family(Family::Monospace).color(col), Shaping::Advanced, None); lab.shape_until_scroll(fs, false);
        for r in lab.layout_runs() { for g in r.glyphs { let pg = g.physical((x, y + r.line_y), 1.0); if let Some(im) = sc.get_image(fs, pg.cache_key) { let mut p = Pixmap::new(im.placement.width.max(1), im.placement.height.max(1)).unwrap(); let (cr, cg, cb, ca) = (col.r(), col.g(), col.b(), col.a()); for (idx, &al) in im.data.iter().enumerate() { let af = (al as f32 / 255.0) * (ca as f32 / 255.0); p.pixels_mut()[idx] = ColorU8::from_rgba((cr as f32 * af) as u8, (cg as f32 * af) as u8, (cb as f32 * af) as u8, (255.0 * af) as u8).premultiply(); } pix.draw_pixmap(pg.x + im.placement.left, pg.y - im.placement.top, p.as_ref(), &PixmapPaint::default(), Transform::identity(), None); } } }
    }
}

impl ApplicationHandler for App {
    fn resumed(&mut self, event_loop: &winit::event_loop::ActiveEventLoop) {
        let window = Arc::new(event_loop.create_window(WindowAttributes::default().with_title("Vybe IDE").with_inner_size(winit::dpi::LogicalSize::new(1200.0, 900.0))).unwrap());
        let ctx = Context::new(window.clone()).unwrap(); let surf = Surface::new(&ctx, window.clone()).unwrap(); let sz = window.inner_size();
        let lang = load_language("rust").expect("load rust");
        let my_editor = MyEditor::from_text("// Welcome to Vybe IDE\nfn main() {\n    println!(\"Multi-file support active!\");\n}", &lang);
        let uri = "file:///Users/youness/www/html/vybe/welcome.rs".to_string();
        let widget = CodeEditorWidget::new(my_editor, &mut self.font_system, self.lsp.clone(), uri);
        self.tabs.push(Tab { name: "welcome.rs".to_string(), path: None, widget, is_sticky: true, buffer: None, is_modified: false });
        self.window = Some(window); self.context = Some(ctx); self.surface = Some(surf);
        self.pixmap = Some(Pixmap::new(sz.width, sz.height).unwrap()); self.surface.as_mut().unwrap().resize(NonZeroU32::new(sz.width).unwrap(), NonZeroU32::new(sz.height).unwrap()).unwrap();
    }
    fn window_event(&mut self, event_loop: &winit::event_loop::ActiveEventLoop, _id: winit::window::WindowId, event: WindowEvent) {
        match event {
            WindowEvent::CloseRequested => event_loop.exit(), WindowEvent::ModifiersChanged(m) => self.modifiers = m,
            WindowEvent::Resized(sz) => { if let (Some(s), Some(w)) = (&mut self.surface, &self.window) { if sz.width > 0 && sz.height > 0 { s.resize(NonZeroU32::new(sz.width).unwrap(), NonZeroU32::new(sz.height).unwrap()).expect("resize surface"); self.pixmap = Some(Pixmap::new(sz.width, sz.height).unwrap()); w.request_redraw(); } } }
            WindowEvent::MouseWheel { delta, .. } => { 
                let a = match delta { 
                    MouseScrollDelta::LineDelta(_, y) => y * 120.0, 
                    MouseScrollDelta::PixelDelta(pos) => pos.y as f32 * 2.0 
                }; 
                if self.mouse_pos.1 / SCALE < TAB_BAR_HEIGHT {
                    self.tab_scroll_x -= a;
                } else if self.active_tab < self.tabs.len() {
                    self.tabs[self.active_tab].widget.scroll_y -= a;
                }
                self.window.as_ref().unwrap().request_redraw(); 
            }
            WindowEvent::KeyboardInput { event, .. } => {
                if event.state == ElementState::Pressed {
                    if self.tabs.is_empty() { return; }
                    let mut acted = true; 
                    let _theme = self.tabs[self.active_tab].widget.theme;
                    let tab = &mut self.tabs[self.active_tab];
                    let w = &mut tab.widget;
                    let cmd = self.modifiers.state().super_key() || self.modifiers.state().control_key();
                    let alt = self.modifiers.state().alt_key(); let shift = self.modifiers.state().shift_key();

                    let key_str = match event.key_without_modifiers() {
                        Key::Character(c) => c.to_lowercase(),
                        Key::Named(nk) => format!("{:?}", nk),
                        _ => String::new(),
                    };

                    for kb in &self.keybindings {
                        if kb.key == key_str && kb.cmd == cmd && kb.shift == shift && kb.alt == alt {
                            match kb.action.as_str() {
                                "Undo" => {
                                    if let Some((text, line, col)) = w.my_editor.undo(w.editor.cursor().line, w.editor.cursor().index) {
                                        w.editor.with_buffer_mut(|b| b.set_text(&mut self.font_system, &text, &Attrs::new().family(Family::Monospace), Shaping::Advanced, None));
                                        w.editor.set_cursor(Cursor::new(line, col));
                                    }
                                }
                                "Redo" => {
                                    if let Some((text, line, col)) = w.my_editor.redo(w.editor.cursor().line, w.editor.cursor().index) {
                                        w.editor.with_buffer_mut(|b| b.set_text(&mut self.font_system, &text, &Attrs::new().family(Family::Monospace), Shaping::Advanced, None));
                                        w.editor.set_cursor(Cursor::new(line, col));
                                    }
                                }
                                "SelectAll" => {
                                    w.editor.set_cursor(Cursor::new(0, 0));
                                    let last_line = w.editor.with_buffer(|b| b.lines.len() - 1);
                                    let last_col = w.editor.with_buffer(|b| b.lines[last_line].text().len());
                                    w.editor.set_selection(Selection::Normal(Cursor::new(last_line, last_col)));
                                }
                                "MoveBufferStart" => w.editor.action(&mut self.font_system, Action::Motion(Motion::BufferStart)),
                                "MoveBufferEnd" => w.editor.action(&mut self.font_system, Action::Motion(Motion::BufferEnd)),
                                "MoveLineStart" => w.editor.action(&mut self.font_system, Action::Motion(Motion::Home)),
                                "MoveLineEnd" => w.editor.action(&mut self.font_system, Action::Motion(Motion::End)),
                                "MoveWordLeft" => w.editor.action(&mut self.font_system, Action::Motion(Motion::LeftWord)),
                                "MoveWordRight" => w.editor.action(&mut self.font_system, Action::Motion(Motion::RightWord)),
                                "Save" => { println!("Saving document: {}", tab.name); tab.is_modified = false; }
                                "Find" => { w.is_search_open = true; if let Some(t) = w.editor.copy_selection() { if !t.is_empty() { w.search_query = t; } } else { w.search_query.clear(); } }
                                "Replace" => { w.is_search_open = true; w.is_replace_open = !w.is_replace_open; }
                                _ => { acted = false; }
                            }
                            if acted { w.needs_reshape = true; w.sync(); self.window.as_ref().unwrap().request_redraw(); return; }
                        }
                    }

                    match event.key_without_modifiers() {
                        Key::Character(c) if cmd && (c == "=" || c == "+") => { w.set_zoom(&mut self.font_system, 1.0); }
                        Key::Character(c) if cmd && c == "-" => { w.set_zoom(&mut self.font_system, -1.0); }
                        Key::Character(c) if cmd && c == "0" => { w.font_size = 14.0; w.set_zoom(&mut self.font_system, 0.0); }
                        Key::Character(c) if cmd && (c == "w" || c == "W") => { w.show_whitespace = !w.show_whitespace; }
                        Key::Character(c) if alt && (c == "z" || c == "Z") => { w.wrap_lines = !w.wrap_lines; w.needs_reshape = true; }
                        Key::Character(c) if cmd && (c == "m" || c == "M") => { if let Some(p) = w.my_editor.find_matching_bracket(w.editor.cursor().line, w.editor.cursor().index, &w.lang_def) { w.editor.set_cursor(Cursor::new(p.0, p.1)); } }
                        Key::Character(c) if cmd && (c == "p" || c == "P") => { self.is_quick_open = !self.is_quick_open; self.quick_open_query.clear(); }
                         Key::Named(NamedKey::Home) => {
                              let cli = w.editor.cursor().line; let cur = w.editor.cursor().index;
                              let line_text = w.editor.with_buffer(|b| b.lines[cli].text().to_string());
                              let first_byte_idx = line_text.char_indices().find(|&(_, c)| !c.is_whitespace()).map(|(i, _)| i).unwrap_or(line_text.len());
                              if cur == first_byte_idx { w.editor.action(&mut self.font_system, Action::Motion(Motion::Home)); }
                              else { w.editor.set_cursor(Cursor::new(cli, first_byte_idx)); }
                         }
                        Key::Named(NamedKey::End) => w.editor.action(&mut self.font_system, Action::Motion(Motion::End)),
                        Key::Character(c) if cmd && shift && (c == "k" || c == "K") => { w.editor.action(&mut self.font_system, Action::Motion(Motion::End)); w.editor.action(&mut self.font_system, Action::Backspace); w.editor.action(&mut self.font_system, Action::Motion(Motion::Home)); let len = w.editor.with_buffer(|b| b.lines[w.editor.cursor().line].text().len()); for _ in 0..len { w.editor.action(&mut self.font_system, Action::Delete); } w.editor.action(&mut self.font_system, Action::Delete); }
                        Key::Named(NamedKey::Backspace) => if w.is_search_open { if w.is_replace_open && alt { w.replace_query.pop(); } else { w.search_query.pop(); } } else { w.my_editor.save_snapshot(w.editor.cursor().line, w.editor.cursor().index); w.editor.action(&mut self.font_system, Action::Backspace); tab.is_modified = true; }
                        Key::Named(NamedKey::Delete) => { w.my_editor.save_snapshot(w.editor.cursor().line, w.editor.cursor().index); w.editor.action(&mut self.font_system, Action::Delete); tab.is_modified = true; }
                        Key::Named(NamedKey::Enter) => {
                            if self.is_quick_open {
                                self.is_quick_open = false;
                                // In a real IDE, we'd fuzzy search and pick. For win, just close.
                            } else if w.is_search_open { w.find_next(&mut self.font_system); } 
                            else { 
                                w.my_editor.save_snapshot(w.editor.cursor().line, w.editor.cursor().index);
                                let line_idx = w.editor.cursor().line;
                                let byte_off = w.editor.with_buffer(|b| {
                                    let mut total = 0;
                                    for i in 0..line_idx { total += b.lines[i].text().len() + 1; }
                                    total + w.editor.cursor().index
                                });
                                w.my_editor.insert_newline(byte_off, &w.lang_def);
                                w.editor.action(&mut self.font_system, Action::Enter); // ensure cosmic-text syncs too
                                w.needs_reshape = true; w.sync(); tab.is_modified = true;
                            }
                        }
                        Key::Named(NamedKey::Escape) => { self.is_quick_open = false; w.is_search_open = false; w.context_menu = None; }
                        Key::Character(c) if cmd && (c == "c" || c == "C") => { if let Some(t) = w.editor.copy_selection() { if let Some(cb) = &mut self.clipboard { let _ = cb.set_text(t); } } }
                        Key::Character(c) if cmd && (c == "v" || c == "V") => { 
                             if let Some(cb) = &mut self.clipboard { 
                                 if let Ok(t) = cb.get_text() { 
                                     w.my_editor.save_snapshot(w.editor.cursor().line, w.editor.cursor().index);
                                     
                                     // Handle selection replacement
                                     if let Some((start, end)) = w.editor.selection_bounds() {
                                         let b = |c: Cursor, ed: &CodeEditorWidget| ed.editor.with_buffer(|buf| {
                                             let mut total = 0;
                                             // cosmic-text 0.18 cursor indices are reliable. 
                                             for i in 0..c.line { total += buf.lines[i].text().len() + 1; }
                                             total + c.index
                                         });
                                         let s_off = b(start, w);
                                         let e_off = b(end, w);
                                         w.my_editor.delete_range(s_off.min(e_off), s_off.max(e_off), &w.lang_def);
                                         w.editor.action(&mut self.font_system, Action::Delete);
                                     }

                                     let byte_off = w.editor.with_buffer(|b| {
                                         let cli = w.editor.cursor().line;
                                         let mut total = 0;
                                         for i in 0..cli { total += b.lines[i].text().len() + 1; }
                                         total + w.editor.cursor().index
                                     });
                                     
                                     let (new_line, new_col) = w.my_editor.insert_string(byte_off, &t, &w.lang_def);
                                     
                                     // Sync cosmic-text
                                     w.editor.with_buffer_mut(|b| {
                                         b.set_text(&mut self.font_system, &w.my_editor.rope().to_string(), &Attrs::new().family(Family::Monospace), Shaping::Advanced, None);
                                     });
                                     w.editor.set_cursor(Cursor::new(new_line, new_col));
                                     tab.is_modified = true; 
                                 } 
                             } 
                        }
                        Key::Character(c) if cmd && (c == "x" || c == "X") => { if let Some(t) = w.editor.copy_selection() { w.my_editor.save_snapshot(w.editor.cursor().line, w.editor.cursor().index); if let Some(cb) = &mut self.clipboard { let _ = cb.set_text(t); } w.editor.action(&mut self.font_system, Action::Delete); tab.is_modified = true; } }
                        Key::Character(c) if cmd && (c == "d" || c == "D") => {
                             w.my_editor.save_snapshot(w.editor.cursor().line, w.editor.cursor().index);
                             if w.editor.selection_bounds().is_none() {
                                 let li = w.editor.cursor().line;
                                 w.my_editor.duplicate_line(li);
                             } else if let Some(t) = w.editor.copy_selection() {
                                 w.find_next(&mut self.font_system); // Simplified "Next occurrence" logic
                             }
                             w.needs_reshape = true; tab.is_modified = true;
                        }
                        Key::Named(NamedKey::ArrowUp) if alt => { w.my_editor.save_snapshot(w.editor.cursor().line, w.editor.cursor().index); let li = w.editor.cursor().line; w.my_editor.move_line_up(li); w.needs_reshape = true; tab.is_modified = true; }
                        Key::Named(NamedKey::ArrowDown) if alt => { w.my_editor.save_snapshot(w.editor.cursor().line, w.editor.cursor().index); let li = w.editor.cursor().line; if shift { w.my_editor.duplicate_line(li); } else { w.my_editor.move_line_down(li); } w.needs_reshape = true; tab.is_modified = true; }
                        Key::Named(NamedKey::ArrowLeft) => {
                            if shift && w.editor.selection_bounds().is_none() { w.editor.set_selection(Selection::Normal(w.editor.cursor())); }
                            w.editor.action(&mut self.font_system, Action::Motion(Motion::Left));
                            if !shift { w.editor.set_selection(Selection::None); }
                        }
                        Key::Named(NamedKey::ArrowRight) => {
                            if shift && w.editor.selection_bounds().is_none() { w.editor.set_selection(Selection::Normal(w.editor.cursor())); }
                            w.editor.action(&mut self.font_system, Action::Motion(Motion::Right));
                            if !shift { w.editor.set_selection(Selection::None); }
                        }
                        Key::Named(NamedKey::ArrowUp) => {
                            if shift && w.editor.selection_bounds().is_none() { w.editor.set_selection(Selection::Normal(w.editor.cursor())); }
                            w.editor.action(&mut self.font_system, Action::Motion(Motion::Up));
                            if !shift { w.editor.set_selection(Selection::None); }
                        }
                        Key::Named(NamedKey::ArrowDown) => {
                            if shift && w.editor.selection_bounds().is_none() { w.editor.set_selection(Selection::Normal(w.editor.cursor())); }
                            w.editor.action(&mut self.font_system, Action::Motion(Motion::Down));
                            if !shift { w.editor.set_selection(Selection::None); }
                        }
                        Key::Character(c) if cmd && c == "z" => { 
                            if shift {
                                if let Some((text, line, col)) = w.my_editor.redo(w.editor.cursor().line, w.editor.cursor().index) {
                                    w.editor.with_buffer_mut(|b| b.set_text(&mut self.font_system, &text, &Attrs::new().family(Family::Monospace), Shaping::Advanced, None));
                                    let safe_line = line.min(w.editor.with_buffer(|b| b.lines.len().saturating_sub(1)));
                                    let safe_col = w.editor.with_buffer(|b| if safe_line < b.lines.len() { col.min(b.lines[safe_line].text().len()) } else { 0 });
                                    w.editor.set_cursor(Cursor::new(safe_line, safe_col));
                                }
                            } else {
                                if let Some((text, line, col)) = w.my_editor.undo(w.editor.cursor().line, w.editor.cursor().index) {
                                    w.editor.with_buffer_mut(|b| b.set_text(&mut self.font_system, &text, &Attrs::new().family(Family::Monospace), Shaping::Advanced, None));
                                    let safe_line = line.min(w.editor.with_buffer(|b| b.lines.len().saturating_sub(1)));
                                    let safe_col = w.editor.with_buffer(|b| if safe_line < b.lines.len() { col.min(b.lines[safe_line].text().len()) } else { 0 });
                                    w.editor.set_cursor(Cursor::new(safe_line, safe_col));
                                }
                            }
                        }
                        Key::Character(c) if cmd && c == "a" => {
                            w.editor.set_cursor(Cursor::new(0, 0));
                            let last_line = w.editor.with_buffer(|b| b.lines.len() - 1);
                            let last_col = w.editor.with_buffer(|b| b.lines[last_line].text().len());
                            w.editor.set_selection(Selection::Normal(Cursor::new(last_line, last_col)));
                        }
                        _ => { if let Some(t) = event.text { if !cmd {
                            w.my_editor.save_snapshot(w.editor.cursor().line, w.editor.cursor().index);
                            
                            // Handling selection replacement on type
                            if let Some((start, end)) = w.editor.selection_bounds() {
                                 let b = |c: Cursor, ed: &CodeEditorWidget| ed.editor.with_buffer(|buf| {
                                    let mut total = 0;
                                    for i in 0..c.line { total += buf.lines[i].text().len() + 1; }
                                    total + c.index
                                });
                                let s_off = b(start, w);
                                let e_off = b(end, w);
                                w.my_editor.delete_range(s_off.min(e_off), s_off.max(e_off), &w.lang_def);
                                // cosmic-text Action::Insert handles UI deletion
                            }

                            for ch in t.chars() { if !ch.is_control() || ch == '\t' || ch == '\n' { if w.is_search_open { if w.is_replace_open && alt { w.replace_query.push(ch); } else { w.search_query.pop(); w.search_query.push(ch); } } else { 
                             let mut skip = false; if let Some(cl) = match ch { ')'=>Some(')'),'}'=>Some('}'),']'=>Some(']'),'"'=>Some('"'),'\''=>Some('\''),_=>None } { 
                                 let cli = w.editor.cursor().line; let cur = w.editor.cursor().index; 
                                 let line_text = w.editor.with_buffer(|b| b.lines[cli].text().to_string());
                                 let next_ch = line_text[cur..].chars().next();
                                 if next_ch == Some(cl) { w.editor.action(&mut self.font_system, Action::Motion(Motion::Right)); skip = true; } 
                             }
                            if !skip { w.editor.action(&mut self.font_system, Action::Insert(ch)); tab.is_modified = true; if let Some(cl) = match ch { '('=>Some(')'),'{'=>Some('}'),'['=>Some(']'),'"'=>Some('"'),'\''=>Some('\''),_=>None } { w.editor.action(&mut self.font_system, Action::Insert(cl)); w.editor.action(&mut self.font_system, Action::Motion(Motion::Left)); } }
                        } } } } else { acted = false; } } else { acted = false; } }
                    }
                    if acted { 
                        let w = &mut self.tabs[self.active_tab].widget; 
                        w.needs_reshape = true; 
                        w.sync(); 
                        self.pending_lsp_update = true;
                        self.last_lsp_update = Instant::now();
                        self.window.as_ref().unwrap().request_redraw(); 
                    }
                }
            }
            WindowEvent::CursorMoved { position, .. } => { 
                self.mouse_pos = (position.x as f32, position.y as f32); 
                let mx = self.mouse_pos.0 / SCALE;
                let my = self.mouse_pos.1 / SCALE;
                let height = self.pixmap.as_ref().unwrap().height() as f32 / SCALE;

                // 1. Splitter Hover/Drag
                let split_start = self.explorer_width;
                let split_end = self.explorer_width + SPLITTER_WIDTH;
                let was_hovering_splitter = self.hovering_splitter;
                self.hovering_splitter = mx >= split_start && mx <= split_end;

                if self.is_dragging_splitter {
                    self.explorer_width = (mx - SPLITTER_WIDTH / 2.0).max(50.0).min(600.0);
                }

                if self.hovering_splitter || self.is_dragging_splitter {
                    self.window.as_ref().unwrap().set_cursor(winit::window::CursorIcon::ColResize);
                } else {
                    self.window.as_ref().unwrap().set_cursor(winit::window::CursorIcon::Default);
                }

                if let Some(mut dropdown) = self.lang_dropdown.take() {
                    let (w, h) = dropdown.get_size();
                    let menu_x = (self.pixmap.as_ref().unwrap().width() as f32 / SCALE - w - 20.0).max(10.0);
                    let menu_y = (height - FOOTER_HEIGHT - h - 10.0).max(10.0);
                    dropdown.handle_mouse(mx, my, menu_x, menu_y, false);
                    self.lang_dropdown = Some(dropdown);
                }
                
                if let Some(mut dropdown) = self.theme_dropdown.take() {
                    let (w, h) = dropdown.get_size();
                    let theme_label = format!("Theme: {}", self.get_theme_name());
                    let lang_label = format!("Language: {}", self.current_lang);
                    let label_x = (self.pixmap.as_ref().unwrap().width() as f32 / SCALE) - (lang_label.len() as f32 * 9.0 + 20.0);
                    let theme_x = label_x - (theme_label.len() as f32 * 9.0 + 30.0);
                    let menu_x = theme_x.min(self.pixmap.as_ref().unwrap().width() as f32 / SCALE - w - 10.0).max(10.0);
                    let menu_y = (height - FOOTER_HEIGHT - h - 10.0).max(10.0);
                    dropdown.handle_mouse(mx, my, menu_x, menu_y, false);
                    self.theme_dropdown = Some(dropdown);
                }

                if self.lang_dropdown.is_some() || self.theme_dropdown.is_some() {
                    self.window.as_ref().unwrap().request_redraw();
                }

                // 2. Tab Close Hover
                let last_tab_hover = self.hovering_tab_close;
                self.hovering_tab_close = None;
                let ed_start_x = self.explorer_width + SPLITTER_WIDTH + 1.0;
                if my < TAB_BAR_HEIGHT && mx > ed_start_x {
                    let mut tx = ed_start_x;
                    for i in 0..self.tabs.len() {
                        let tw = 160.0;
                        if mx >= tx + tw - 30.0 && mx <= tx + tw - 5.0 {
                            self.hovering_tab_close = Some(i);
                            break;
                        }
                        tx += tw;
                    }
                }

                // 3. Editor Drag
                let mut needs_editor_redraw = false;
                if self.is_dragging && !self.is_dragging_splitter && self.active_tab < self.tabs.len() { 
                    let ed_top = TAB_BAR_HEIGHT * SCALE;
                    let r = Rect::from_xywh(ed_start_x * SCALE, ed_top, self.pixmap.as_ref().unwrap().width() as f32 - ed_start_x * SCALE, self.pixmap.as_ref().unwrap().height() as f32 - (ed_top + FOOTER_HEIGHT * SCALE)).unwrap(); 
                    self.tabs[self.active_tab].widget.handle_mouse(&mut self.font_system, self.mouse_pos.0, self.mouse_pos.1, r, None, &mut self.clipboard); 
                    needs_editor_redraw = true;
                } 

                // Smart Redraw: Only if interaction state changed or dragging
                if was_hovering_splitter != self.hovering_splitter || 
                   last_tab_hover != self.hovering_tab_close || 
                   self.is_dragging_splitter || 
                   self.lang_dropdown.is_some() ||
                   needs_editor_redraw {
                    self.window.as_ref().unwrap().request_redraw();
                }
            }
            WindowEvent::Focused(false) | WindowEvent::CursorLeft { .. } => {
                self.is_dragging = false;
                self.is_dragging_splitter = false;
            }
                WindowEvent::MouseInput { state, button, .. } => {
                    let mx = self.mouse_pos.0 / SCALE;
                    let my = self.mouse_pos.1 / SCALE;
                    let pw = self.pixmap.as_ref().unwrap().width() as f32;
                    let ph = self.pixmap.as_ref().unwrap().height() as f32 / SCALE;
                    let height = ph; // restore height alias

                    if state == ElementState::Pressed && button == MouseButton::Right {
                        self.tabs[self.active_tab].widget.context_menu = Some(((mx, my), vec!["Cut".into(), "Copy".into(), "Paste".into(), "Go to Def".into()]));
                        self.window.as_ref().unwrap().request_redraw();
                        return;
                    }

                    if state == ElementState::Pressed && button == MouseButton::Left {
                        // 0a. Language Picker Menu Intercept
                        if let Some(mut dropdown) = self.lang_dropdown.take() {
                            let (w, h) = dropdown.get_size();
                            let label_x = (pw / SCALE) - (format!("Language: {}", self.current_lang).len() as f32 * 9.0 + 20.0);
                            let menu_x = label_x.min(pw / SCALE - w - 10.0).max(10.0);
                            let menu_y = (height - FOOTER_HEIGHT - h - 10.0).max(10.0);
                            
                            match dropdown.handle_mouse(mx, my, menu_x, menu_y, true) {
                                DropdownEvent::Selected(idx) => {
                                    if let Some(new_lang) = self.all_languages.get(idx).cloned() {
                                        self.current_lang = new_lang.clone();
                                        let tab = &mut self.tabs[self.active_tab];
                                        let uri = tab.path.clone().unwrap_or_else(|| format!("file:///Users/youness/www/html/vybe/{}", tab.name));
                                        tab.widget.set_language(&new_lang, uri);
                                    }
                                    self.lang_dropdown = None;
                                }
                                DropdownEvent::Closed => { self.lang_dropdown = None; }
                                DropdownEvent::None => self.lang_dropdown = Some(dropdown),
                                _ => {}
                            }
                            self.window.as_ref().unwrap().request_redraw(); return;
                        }

                        // 0b. Theme Picker Menu Intercept
                        if let Some(mut dropdown) = self.theme_dropdown.take() {
                            let (w, h) = dropdown.get_size();
                            let theme_label = format!("Theme: {}", self.get_theme_name());
                            let lang_label = format!("Language: {}", self.current_lang);
                            let label_x = (pw / SCALE) - (lang_label.len() as f32 * 9.0 + 20.0);
                            let theme_x = label_x - (theme_label.len() as f32 * 9.0 + 30.0);
                            let menu_x = theme_x.min(pw / SCALE - w - 10.0).max(10.0);
                            let menu_y = (height - FOOTER_HEIGHT - h - 10.0).max(10.0);
                            
                            match dropdown.handle_mouse(mx, my, menu_x, menu_y, true) {
                                DropdownEvent::Selected(idx) => {
                                    self.current_theme_idx = idx;
                                    let new_theme = self.active_theme();
                                    for tab in &mut self.tabs { tab.widget.theme = new_theme; tab.widget.needs_reshape = true; }
                                    self.window.as_ref().unwrap().request_redraw();
                                    return;
                                }
                                DropdownEvent::None => self.theme_dropdown = Some(dropdown),
                                _ => {}
                            }
                            self.window.as_ref().unwrap().request_redraw(); return;
                        }

                        // 1. Minimap Hit-testing
                        if mx > pw / SCALE - MINIMAP_WIDTH {
                            if self.active_tab < self.tabs.len() {
                                let tab = &mut self.tabs[self.active_tab];
                                let mut th = 0.0; tab.widget.editor.with_buffer(|b| { for r in b.layout_runs() { if !tab.widget.is_line_hidden(r.line_i) { th += r.line_height; } } });
                                let mry = (my - TAB_BAR_HEIGHT) / (height - TAB_BAR_HEIGHT - FOOTER_HEIGHT);
                                tab.widget.scroll_y = (mry * th).max(0.0);
                                self.window.as_ref().unwrap().request_redraw();
                                return;
                            }
                        }


                        // 3. Status Bar Click
                        if my >= height - FOOTER_HEIGHT {
                            // Breadcrumb segments hit-testing
                            for (rect, path) in &self.breadcrumb_rects {
                                if mx * SCALE >= rect.left() && mx * SCALE <= rect.right() && my * SCALE >= rect.top() && my * SCALE <= rect.bottom() {
                                    println!("Revealing in explorer: {}", path);
                                    continue;
                                }
                            }

                            let lang_label = format!("Language: {}", self.current_lang);
                            let theme_label = format!("Theme: {}", self.get_theme_name());
                            let label_x = (pw / SCALE) - (lang_label.len() as f32 * 9.0 + 20.0);
                            let theme_x = label_x - (theme_label.len() as f32 * 9.0 + 30.0);

                            if mx >= label_x {
                                let active_idx = self.all_languages.iter().position(|l| l == &self.current_lang).unwrap_or(0);
                                self.lang_dropdown = Some(Dropdown::new(self.all_languages.clone(), active_idx, SCALE, None));
                            } else if mx >= theme_x && mx < label_x {
                                let theme_names = vec![
                                    "Silicon Green".into(), "Cloud Blue".into(), "Coffee Cream".into(), "Sakura Pink".into(), 
                                    "One Dark".into(), "Monokai".into(), "GitHub Light".into(), "Solarized Light".into(), 
                                    "Midnight".into(), "Aura".into(), "Veridian".into(), "Rose".into(),
                                    "Cyber".into(), "Titanium".into(), "Indigo Night".into()
                                ];
                                let mut d = Dropdown::new(theme_names, self.current_theme_idx, SCALE, None);
                                d.num_cols = 2; d.col_w = 160.0; // Balanced 2-column grid for 10 vivid presets
                                self.theme_dropdown = Some(d);
                            }
                            self.window.as_ref().unwrap().request_redraw(); return;
                        }

                        // 4. Tab Bar Click
                    let ed_start_x = self.explorer_width + SPLITTER_WIDTH + 1.0;
                    if my < TAB_BAR_HEIGHT && mx > ed_start_x {
                        if let Some(idx) = self.hovering_tab_close {
                             println!("DEBUG: Tab Bar Click -> Close Tab {}", idx);
                             self.tabs.remove(idx);
                             if self.tabs.is_empty() { self.active_tab = 0; }
                             else if self.active_tab >= self.tabs.len() { self.active_tab = self.tabs.len() - 1; }
                             self.window.as_ref().unwrap().request_redraw(); return;
                        }
                        let tab_idx = ((mx - ed_start_x) / 160.0) as usize;
                        if tab_idx < self.tabs.len() { 
                            println!("DEBUG: Tab Bar Click -> Select Tab {}", tab_idx);
                            self.active_tab = tab_idx; self.window.as_ref().unwrap().request_redraw(); 
                        }
                        return;
                    }

                    // 3. Sidebar Click
                    if mx < self.explorer_width {
                        println!("DEBUG: Sidebar Click at mx={}, my={}", mx, my);
                        let now = Instant::now();
                        let is_double = (now - self.last_click_time) < Duration::from_millis(300);
                        self.last_click_time = now;
                        match self.tree_view.handle_mouse(self.mouse_pos.0, self.mouse_pos.1, 0.0, TAB_BAR_HEIGHT * SCALE) {
                             TreeEvent::Open(path) => {
                                 if let Some(idx) = self.tabs.iter().position(|t| t.path.as_ref() == Some(&path)) {
                                     println!("DEBUG: Sidebar -> Reveal existing tab: {}", path);
                                     self.active_tab = idx; if is_double { self.tabs[idx].is_sticky = true; }
                                     self.window.as_ref().unwrap().request_redraw(); return;
                                 }
                                 if let Ok(content) = fs::read_to_string(&path) {
                                     println!("DEBUG: Sidebar -> Open new tab: {}", path);
                                     let ext = path.split('.').last().unwrap_or("txt");
                                     let lang_name = match ext { "rs" => "rust", "js" => "javascript", "bas" | "vb" => "vb", "cs" => "csharp", _ => "text" };
                                     let lang = load_language(lang_name).or_else(|| load_language("rust")).expect("rust language not found");
                                     let my_editor = MyEditor::from_text(&content, &lang);
                                     let uri = format!("file://{}", path);
                                     let mut widget = CodeEditorWidget::new(my_editor, &mut self.font_system, self.lsp.clone(), uri.clone());
                                     widget.set_language(lang_name, uri);
                                     let name = Path::new(&path).file_name().unwrap_or_default().to_string_lossy().to_string();
                                     let new_tab = Tab { name, path: Some(path.clone()), widget, is_sticky: is_double, buffer: None, is_modified: false };
                                     self.tabs.push(new_tab); self.active_tab = self.tabs.len() - 1; self.tree_view.reveal_path(&path);
                                 }
                             }
                             _ => { self.window.as_ref().unwrap().request_redraw(); }
                        }
                        self.window.as_ref().unwrap().request_redraw(); return;
                    }

                    // 4. Splitter Click
                    if self.hovering_splitter {
                        println!("DEBUG: Splitter Click -> Start Resizing");
                        self.is_dragging_splitter = true; self.window.as_ref().unwrap().request_redraw(); return;
                    }

                    // 5. Editor Click (Deep Isolation)
                    if !self.tabs.is_empty() {
                        let ed_top = TAB_BAR_HEIGHT * SCALE;
                        let ed_bottom = height * SCALE - FOOTER_HEIGHT * SCALE;
                        if self.mouse_pos.1 >= ed_top && self.mouse_pos.1 < ed_bottom && mx >= self.explorer_width + SPLITTER_WIDTH {
                            println!("DEBUG: Editor Click at mx={}, my={}", mx, my);
                            let rect = Rect::from_xywh(ed_start_x * SCALE, ed_top, self.pixmap.as_ref().unwrap().width() as f32 - ed_start_x * SCALE, ed_bottom - ed_top).unwrap();
                            self.click_count = if Instant::now().duration_since(self.last_click_time) < Duration::from_millis(500) { (self.click_count % 3) + 1 } else { 1 }; self.last_click_time = Instant::now();
                            self.tabs[self.active_tab].widget.handle_mouse(&mut self.font_system, self.mouse_pos.0, self.mouse_pos.1, rect, Some((self.click_count, button, self.modifiers)), &mut self.clipboard);
                            self.is_dragging = true; self.window.as_ref().unwrap().request_redraw();
                        }
                    }
                } else if state == ElementState::Released {
                    println!("DEBUG: Mouse Released");
                    self.is_dragging = false; 
                    self.is_dragging_splitter = false;
                    self.window.as_ref().unwrap().request_redraw();
                }
            }
            WindowEvent::RedrawRequested => self.render(),
            _ => {}
        }
    }
}

pub fn run_gui(my_editor: MyEditor) {
    let el = EventLoop::new().expect("event loop"); el.set_control_flow(ControlFlow::Wait);
    let mut app = App::new(my_editor); el.run_app(&mut app).expect("run app");
}
