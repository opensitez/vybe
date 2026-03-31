//! Proportional layout system for the IDE.
//!
//! All positions are in logical pixels. Scale is only applied at render time.

/// Rectangle in logical pixels.
#[derive(Clone, Copy, Debug)]
pub struct Rect {
    pub x: f32,
    pub y: f32,
    pub w: f32,
    pub h: f32,
}

impl Rect {
    pub fn contains(&self, px: f32, py: f32) -> bool {
        px >= self.x && px < self.x + self.w && py >= self.y && py < self.y + self.h
    }
}

/// Computed layout for every panel in the IDE.
#[derive(Clone, Debug)]
pub struct IdeLayout {
    pub menu_bar: Rect,
    pub toolbar: Rect,
    pub project_explorer: Rect,
    pub splitter: Rect,
    pub toolbox: Rect,
    pub center: Rect,
    pub properties: Rect,
    pub status_bar: Rect,
}

/// Configuration flags and splitter position.
pub struct LayoutConfig {
    pub show_project_explorer: bool,
    pub show_toolbox: bool,
    pub show_properties: bool,
    /// Fraction of left column occupied by the project explorer (0.0..1.0).
    /// The rest goes to the toolbox. Default 0.45.
    pub left_split: f32,
}

impl Default for LayoutConfig {
    fn default() -> Self {
        Self {
            show_project_explorer: true,
            show_toolbox: true,
            show_properties: true,
            left_split: 0.45,
        }
    }
}

const MENU_H: f32 = 28.0;
const TOOLBAR_H: f32 = 36.0;
const STATUS_H: f32 = 24.0;
const LEFT_W: f32 = 220.0;
const PROPERTIES_W: f32 = 260.0;
const SPLITTER_H: f32 = 5.0;

impl IdeLayout {
    pub fn compute(win_w: f32, win_h: f32, cfg: &LayoutConfig) -> Self {
        let menu_bar = Rect { x: 0.0, y: 0.0, w: win_w, h: MENU_H };
        let toolbar = Rect { x: 0.0, y: MENU_H, w: win_w, h: TOOLBAR_H };
        let body_top = MENU_H + TOOLBAR_H;
        let body_h = (win_h - body_top - STATUS_H).max(0.0);
        let status_bar = Rect { x: 0.0, y: win_h - STATUS_H, w: win_w, h: STATUS_H };

        let has_left = cfg.show_project_explorer || cfg.show_toolbox;
        let left_w = if has_left { LEFT_W } else { 0.0 };

        let (project_explorer, splitter, toolbox) =
            if cfg.show_project_explorer && cfg.show_toolbox {
                let explorer_h = (body_h * cfg.left_split).floor();
                let splitter_y = body_top + explorer_h;
                let toolbox_h = body_h - explorer_h - SPLITTER_H;
                (
                    Rect { x: 0.0, y: body_top, w: left_w, h: explorer_h },
                    Rect { x: 0.0, y: splitter_y, w: left_w, h: SPLITTER_H },
                    Rect { x: 0.0, y: splitter_y + SPLITTER_H, w: left_w, h: toolbox_h.max(0.0) },
                )
            } else if cfg.show_project_explorer {
                (
                    Rect { x: 0.0, y: body_top, w: left_w, h: body_h },
                    Rect { x: 0.0, y: 0.0, w: 0.0, h: 0.0 },
                    Rect { x: 0.0, y: 0.0, w: 0.0, h: 0.0 },
                )
            } else if cfg.show_toolbox {
                (
                    Rect { x: 0.0, y: 0.0, w: 0.0, h: 0.0 },
                    Rect { x: 0.0, y: 0.0, w: 0.0, h: 0.0 },
                    Rect { x: 0.0, y: body_top, w: left_w, h: body_h },
                )
            } else {
                (
                    Rect { x: 0.0, y: 0.0, w: 0.0, h: 0.0 },
                    Rect { x: 0.0, y: 0.0, w: 0.0, h: 0.0 },
                    Rect { x: 0.0, y: 0.0, w: 0.0, h: 0.0 },
                )
            };

        let right_w = if cfg.show_properties { PROPERTIES_W } else { 0.0 };
        let properties = Rect { x: win_w - right_w, y: body_top, w: right_w, h: body_h };

        let center_x = left_w;
        let center_w = (win_w - left_w - right_w).max(0.0);
        let center = Rect { x: center_x, y: body_top, w: center_w, h: body_h };

        IdeLayout { menu_bar, toolbar, project_explorer, splitter, toolbox, center, properties, status_bar }
    }
}
