//! Standalone code editor widget — renders a full code editor with syntax highlighting,
//! gutter, minimap, search/replace, and folding. No LSP dependency.

use std::collections::HashMap;
use cosmic_text::{Attrs, AttrsList, Buffer, Color, Family, FontSystem, Metrics, Shaping, SwashCache, Action, Cursor, Selection, Edit};
use tiny_skia::{Paint, Pixmap, PixmapPaint, Rect, Transform, ColorU8, Stroke, PathBuilder};
use winit::event::MouseButton;
use arboard::Clipboard;

use crate::text_editor::TextEditor;
use crate::language::LanguageDef;
pub use crate::text_editor::TokenKind;

/// Default layout constants. Can be overridden per-widget via the config fields.
pub const DEFAULT_SCALE: f32 = 2.0;
pub const DEFAULT_GUTTER_WIDTH: f32 = 64.0;
pub const DEFAULT_MINIMAP_WIDTH: f32 = 80.0;
pub const DEFAULT_FOOTER_HEIGHT: f32 = 24.0;

/// Alias for renderer.rs compatibility — uses the widget's own SCALE.
const SCALE: f32 = DEFAULT_SCALE;
const GUTTER_WIDTH: f32 = DEFAULT_GUTTER_WIDTH;
const MINIMAP_WIDTH: f32 = DEFAULT_MINIMAP_WIDTH;
const FOOTER_HEIGHT: f32 = DEFAULT_FOOTER_HEIGHT;

#[derive(Clone)]
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

pub struct CachedGlyph {
    pub pixmap: Pixmap,
    pub left: i32,
    pub top: i32,
}

pub fn apply_highlighting(editor: &mut cosmic_text::Editor<'static>, my_editor: &TextEditor, attrs: &Attrs, lang: &LanguageDef, theme: &Theme) {
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
    pub editor: cosmic_text::Editor<'static>,
    pub my_editor: TextEditor,
    pub lang_def: LanguageDef,
    pub theme: Theme,
    pub metrics: Metrics,
    pub glyph_cache: HashMap<(cosmic_text::CacheKey, Color, bool), CachedGlyph>,
    pub digit_cache: Vec<CachedGlyph>,
    pub needs_reshape: bool,
    pub scroll_y: f32,
    pub search_query: String,
    pub replace_query: String,
    pub is_search_open: bool,
    pub is_replace_open: bool,
    pub case_sensitive: bool,
    pub context_menu: Option<((f32, f32), Vec<String>)>,
    pub font_size: f32,
    pub show_whitespace: bool,
    pub minimap_pixmap: Option<Pixmap>,
    pub minimap_needs_redraw: bool,
    pub wrap_lines: bool,
    pub diagnostics: Vec<crate::text_editor::DiagnosticInfo>,
    pub matching_bracket: Option<usize>,
}

impl CodeEditorWidget {
    pub fn new(mut my_editor: TextEditor, font_system: &mut FontSystem) -> Self {
        let font_size = 14.0;
        let metrics = Metrics::new(font_size, 20.0).scale(SCALE);
        let lang_def = crate::language::load_language("rust").unwrap_or_else(|| LanguageDef {
            keywords: std::collections::HashSet::new(), type_keywords: std::collections::HashSet::new(), constants: std::collections::HashSet::new(), operators: std::collections::HashSet::new(), ignore_case: false, comments: None, brackets: Vec::new(),
        });
        let theme = Theme::dark();
        my_editor.retokenize_all(&lang_def);
        let mut buffer = Buffer::new(font_system, metrics);
        let text = my_editor.rope.to_string();
        buffer.set_text(font_system, &text, &Attrs::new().family(Family::Monospace), Shaping::Advanced, None);

        let mut widget = Self { editor: cosmic_text::Editor::new(buffer), my_editor, lang_def, theme, metrics, glyph_cache: HashMap::new(), digit_cache: Vec::new(), needs_reshape: true, scroll_y: 0.0, search_query: String::new(), replace_query: String::new(), is_search_open: false, is_replace_open: false, case_sensitive: false, context_menu: None, font_size, show_whitespace: false, minimap_pixmap: None, minimap_needs_redraw: true, wrap_lines: false, diagnostics: Vec::new(), matching_bracket: None };
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

    pub fn set_language(&mut self, lang_name: &str) {
        if let Some(lang) = crate::language::load_language(lang_name) {
            self.lang_def = lang;
            self.my_editor.retokenize_all(&self.lang_def);
            self.my_editor.diagnostics.clear();
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

    pub fn is_line_hidden(&self, li: usize) -> bool {
        for (s, e) in &self.my_editor.folds { if self.my_editor.collapsed_starts.contains(s) && li > *s && li <= *e { return true; } } false
    }


    pub fn find_next(&mut self, _fs: &mut FontSystem) {
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

    #[allow(dead_code)]
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

    #[allow(dead_code)]
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
        let cursor_state = self.editor.cursor(); let selection = self.editor.selection_bounds();
        let selected_text = self.editor.copy_selection();
        let partner = self.my_editor.find_matching_bracket(cursor_state.line, cursor_state.index, &self.lang_def).or_else(|| if cursor_state.index > 0 { self.my_editor.find_matching_bracket(cursor_state.line, cursor_state.index - 1, &self.lang_def) } else { None });
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
                    if diag.line == i {
                        let mut g_x = 0.0; let mut g_w = 0.0;
                        for g in glyphs { if g.start >= diag.col_start && g.start < diag.col_end { if g_w == 0.0 { g_x = g.x; } g_w += g.w; } }
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
                        draw_ui_text(pixmap, fs, sc, if col_folded { "+" } else { "-" }, rect.left() + 5.0 * SCALE, cyo + adj_ly - (12.0 * SCALE), if col_folded { theme.kw } else { Color::rgb(0x85, 0x85, 0x85) }); 
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
            let to_skia = |c: Color| tiny_skia::Color::from_rgba8(c.r(), c.g(), c.b(), c.a());
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
             draw_ui_text(pixmap, fs, sc, &format!("Find: {}", self.search_query), rect.right() - 250.0 * SCALE, rect.top() + 14.0 * SCALE, self.theme.text);
             if self.is_replace_open { draw_ui_text(pixmap, fs, sc, &format!("Replace: {}", self.replace_query), rect.right() - 250.0 * SCALE, rect.top() + 44.0 * SCALE, Color::rgb(206, 145, 120)); }
        }
        if let Some((pos, items)) = &self.context_menu {
            let mut mep = Paint::default(); mep.set_color_rgba8(45, 45, 45, 255);
            pixmap.fill_rect(Rect::from_xywh(pos.0, pos.1, 100.0 * SCALE, (items.len() as f32 * 25.0) * SCALE).unwrap(), &mep, Transform::identity(), None);
            for (i, item) in items.iter().enumerate() { draw_ui_text(pixmap, fs, sc, item, pos.0 + 10.0 * SCALE, pos.1 + (i as f32 * 25.0 + 5.0) * SCALE, self.theme.text); }
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
                draw_ui_text(pixmap, fs, sc, item, *mx * SCALE + 10.0 * SCALE, *my * SCALE + (i as f32 * 25.0 + 5.0) * SCALE, Color::rgb(220, 220, 220));
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

/// Draw monospace UI text at physical pixel coordinates (used for gutter, overlays, etc.)
fn draw_ui_text(pix: &mut Pixmap, fs: &mut FontSystem, sc: &mut SwashCache, text: &str, x: f32, y: f32, col: Color) {
    let mut lab = Buffer::new(fs, Metrics::new(14.0, 20.0).scale(SCALE));
    lab.set_text(fs, text, &Attrs::new().family(Family::Monospace).color(col), Shaping::Advanced, None);
    lab.shape_until_scroll(fs, false);
    for r in lab.layout_runs() {
        for g in r.glyphs {
            let pg = g.physical((x, y + r.line_y), 1.0);
            if let Some(img) = sc.get_image(fs, pg.cache_key) {
                if let Some(mut p) = Pixmap::new(img.placement.width.max(1), img.placement.height.max(1)) {
                    let (cr, cg, cb, ca) = (col.r(), col.g(), col.b(), col.a());
                    for (idx, &al) in img.data.iter().enumerate() {
                        let af = (al as f32 / 255.0) * (ca as f32 / 255.0);
                        p.pixels_mut()[idx] = ColorU8::from_rgba((cr as f32 * af) as u8, (cg as f32 * af) as u8, (cb as f32 * af) as u8, (255.0 * af) as u8).premultiply();
                    }
                    pix.draw_pixmap(pg.x + img.placement.left, pg.y - img.placement.top, p.as_ref(), &PixmapPaint::default(), Transform::identity(), None);
                }
            }
        }
    }
}
