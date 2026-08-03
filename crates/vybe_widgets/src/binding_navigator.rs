//! BindingNavigator — a polished record-navigation toolbar.
//!
//! Renders a horizontal strip of navigation buttons with a position counter:
//!
//!   [ |◀ ] [ ◀ ] [ 1 of 42 ] [ ▶ ] [ ▶| ] [ + ] [ − ]
//!
//! Emits `WidgetEvent::Action("nav:<name>:<action>")` where `<action>` is one
//! of: `first`, `prev`, `next`, `last`, `add`, `remove`.

use super::WidgetColors;
use super::layout::{
    CommandValue, KeyEvent, LayoutRect, MouseButton as LayoutMouseButton, MouseEvent,
    MouseEventKind, PanelWidget, RenderContext, WidgetCommand, WidgetEvent, WidgetId };
use cosmic_text::Color as CosmicColor;
use tiny_skia::*;

/// Which navigation button was pressed.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum NavAction {
    First,
    Previous,
    Next,
    Last,
    Add,
    Remove }

impl NavAction {
    pub fn as_str(&self) -> &'static str {
        match self {
            NavAction::First => "first",
            NavAction::Previous => "prev",
            NavAction::Next => "next",
            NavAction::Last => "last",
            NavAction::Add => "add",
            NavAction::Remove => "remove" }
    }
}

/// An individual button in the navigator bar.
struct NavButton {
    label: &'static str,
    action: NavAction,
    rect: LayoutRect,
    hovered: bool,
    pressed: bool }

pub struct BindingNavigator {
    pub name: String,
    pub id: WidgetId,
    pub position: i32,
    pub count: i32,
    pub colors: WidgetColors,
    rect: LayoutRect,
    buttons: Vec<NavButton>,
    counter_rect: LayoutRect,
    pending_events: Vec<WidgetEvent>,
    focused: bool,
    hovered: bool }

impl BindingNavigator {
    pub fn new(name: &str) -> Self {
        Self {
            name: name.to_string(),
            id: WidgetId::next(),
            position: 0,
            count: 0,
            colors: WidgetColors {
                background: (248, 248, 248, 255),
                border: (200, 200, 200, 255),
                foreground: (50, 50, 50, 255),
                accent: (0, 120, 215, 255),
                ..WidgetColors::default()
            },
            rect: LayoutRect::zero(),
            buttons: Vec::new(),
            counter_rect: LayoutRect::zero(),
            pending_events: Vec::new(),
            focused: false,
            hovered: false }
    }

    pub fn with_name(mut self, name: &str) -> Self {
        self.name = name.to_string();
        self
    }

    /// Set the current record position (0-based) and total count.
    pub fn set_position(&mut self, position: i32, count: i32) {
        self.position = position;
        self.count = count;
    }

    fn relayout(&mut self) {
        let r = self.rect;
        if r.w <= 0.0 || r.h <= 0.0 {
            return;
        }

        self.buttons.clear();
        let btn_h = (r.h - 6.0).max(12.0);
        let btn_w = 28.0_f32;
        let counter_w = 80.0_f32;
        let gap = 2.0_f32;
        let pad = 4.0_f32;

        let defs: &[(&str, NavAction)] = &[("⏮", NavAction::First), ("◀", NavAction::Previous)];
        let defs_right: &[(&str, NavAction)] = &[("▶", NavAction::Next), ("⏭", NavAction::Last)];
        let defs_extra: &[(&str, NavAction)] = &[("+", NavAction::Add), ("−", NavAction::Remove)];

        let mut cx = r.x + pad;
        let cy = r.y + (r.h - btn_h) / 2.0;

        // Left nav buttons
        for &(label, action) in defs {
            self.buttons.push(NavButton {
                label,
                action,
                rect: LayoutRect::new(cx, cy, btn_w, btn_h),
                hovered: false,
                pressed: false });
            cx += btn_w + gap;
        }

        // Counter area
        self.counter_rect = LayoutRect::new(cx, cy, counter_w, btn_h);
        cx += counter_w + gap;

        // Right nav buttons
        for &(label, action) in defs_right {
            self.buttons.push(NavButton {
                label,
                action,
                rect: LayoutRect::new(cx, cy, btn_w, btn_h),
                hovered: false,
                pressed: false });
            cx += btn_w + gap;
        }

        // Separator gap
        cx += gap * 2.0;

        // Add/Remove buttons
        for &(label, action) in defs_extra {
            self.buttons.push(NavButton {
                label,
                action,
                rect: LayoutRect::new(cx, cy, btn_w, btn_h),
                hovered: false,
                pressed: false });
            cx += btn_w + gap;
        }
    }

    fn emit(&mut self, action: NavAction) {
        let event_str = format!("nav:{}:{}", self.name, action.as_str());
        self.pending_events.push(WidgetEvent::Action(event_str));
    }
}

impl PanelWidget for BindingNavigator {
    fn name(&self) -> &str {
        &self.name
    }
    fn widget_id(&self) -> WidgetId {
        self.id
    }
    fn focusable(&self) -> bool {
        true
    }
    fn set_focused(&mut self, f: bool) {
        self.focused = f;
    }
    fn hovered(&self) -> bool {
        self.hovered
    }
    fn set_hovered(&mut self, h: bool) {
        self.hovered = h;
    }

    fn set_rect(&mut self, rect: LayoutRect) {
        self.rect = rect;
        self.relayout();
    }

    fn rect(&self) -> LayoutRect {
        self.rect
    }

    fn render(&mut self, ctx: &mut RenderContext) {
        let r = self.rect;
        if r.w <= 0.0 || r.h <= 0.0 {
            return;
        }

        let ts = Transform::from_scale(ctx.scale, ctx.scale);
        let mut paint = Paint::default();
        paint.anti_alias = true;

        // Toolbar background — subtle gradient feel
        let (br, bg, bb, ba) = self.colors.background;
        paint.set_color_rgba8(br, bg, bb, ba);
        if let Some(path) = super::rounded_rect_path(r.x, r.y, r.w, r.h, 3.0) {
            ctx.pixmap
                .fill_path(&path, &paint, FillRule::Winding, ts, None);
        }

        // Toolbar border
        let (bdr, bdg, bdb, bda) = self.colors.border;
        paint.set_color_rgba8(bdr, bdg, bdb, bda);
        let stroke = Stroke {
            width: 1.0,
            ..Stroke::default()
        };
        if let Some(path) = super::rounded_rect_path(r.x, r.y, r.w, r.h, 3.0) {
            ctx.pixmap.stroke_path(&path, &paint, &stroke, ts, None);
        }

        let font_size = 12.0_f32;
        let (fr, fg, fb, _) = self.colors.foreground;

        // Draw buttons
        for btn in &self.buttons {
            let br = btn.rect;

            // Button background
            if btn.pressed {
                paint.set_color_rgba8(180, 210, 240, 255);
            } else if btn.hovered {
                paint.set_color_rgba8(220, 232, 245, 255);
            } else {
                paint.set_color_rgba8(240, 240, 240, 255);
            }
            if let Some(path) = super::rounded_rect_path(br.x, br.y, br.w, br.h, 3.0) {
                ctx.pixmap
                    .fill_path(&path, &paint, FillRule::Winding, ts, None);
            }

            // Button border
            paint.set_color_rgba8(190, 190, 190, 255);
            if let Some(path) = super::rounded_rect_path(br.x, br.y, br.w, br.h, 3.0) {
                ctx.pixmap.stroke_path(&path, &paint, &stroke, ts, None);
            }

            // Button label — centered
            let tw =
                super::ide_text::measure_text(ctx.font_system, btn.label, font_size, ctx.scale);
            let tx = br.x + (br.w - tw) / 2.0;
            let ty = br.y + (br.h - font_size) / 2.0 - 1.0;
            super::ide_text::draw_text(
                ctx.pixmap,
                ctx.font_system,
                ctx.swash_cache,
                btn.label,
                tx,
                ty,
                font_size,
                CosmicColor::rgba(fr, fg, fb, 255),
                ctx.scale,
            );
        }

        // Counter area: white inset with "N of M"
        let cr = self.counter_rect;
        if cr.w > 0.0 {
            paint.set_color_rgba8(255, 255, 255, 255);
            if let Some(path) = super::rounded_rect_path(cr.x, cr.y, cr.w, cr.h, 2.0) {
                ctx.pixmap
                    .fill_path(&path, &paint, FillRule::Winding, ts, None);
            }
            paint.set_color_rgba8(190, 190, 190, 255);
            if let Some(path) = super::rounded_rect_path(cr.x, cr.y, cr.w, cr.h, 2.0) {
                ctx.pixmap.stroke_path(&path, &paint, &stroke, ts, None);
            }
            let display_pos = if self.count > 0 { self.position + 1 } else { 0 };
            let counter_text = format!("{} of {}", display_pos, self.count);
            let tw =
                super::ide_text::measure_text(ctx.font_system, &counter_text, font_size, ctx.scale);
            let tx = cr.x + (cr.w - tw) / 2.0;
            let ty = cr.y + (cr.h - font_size) / 2.0 - 1.0;
            super::ide_text::draw_text(
                ctx.pixmap,
                ctx.font_system,
                ctx.swash_cache,
                &counter_text,
                tx,
                ty,
                font_size,
                CosmicColor::rgba(fr, fg, fb, 255),
                ctx.scale,
            );
        }
    }

    fn handle_mouse(&mut self, event: &MouseEvent) -> bool {
        if !self.rect.contains(event.x, event.y) {
            // Clear hover/press on all buttons
            for btn in &mut self.buttons {
                btn.hovered = false;
                btn.pressed = false;
            }
            return false;
        }

        match event.kind {
            MouseEventKind::Move => {
                for btn in &mut self.buttons {
                    btn.hovered = btn.rect.contains(event.x, event.y);
                }
                true
            }
            MouseEventKind::Press(LayoutMouseButton::Left) => {
                for btn in &mut self.buttons {
                    btn.pressed = btn.rect.contains(event.x, event.y);
                }
                true
            }
            MouseEventKind::Release(LayoutMouseButton::Left) => {
                let mut fired: Option<NavAction> = None;
                for btn in &mut self.buttons {
                    if btn.pressed && btn.rect.contains(event.x, event.y) {
                        fired = Some(btn.action);
                    }
                    btn.pressed = false;
                }
                if let Some(action) = fired {
                    self.emit(action);
                }
                true
            }
            _ => false }
    }

    fn handle_key(&mut self, _event: &KeyEvent) -> bool {
        false
    }

    fn drain_events(&mut self) -> Vec<WidgetEvent> {
        std::mem::take(&mut self.pending_events)
    }

    fn handle_command(&mut self, cmd: &WidgetCommand) -> CommandValue {
        match cmd {
            WidgetCommand::SetValue(v) => {
                // Encode position in lower 32 bits, count in upper bits
                // or just set position directly
                self.position = *v as i32;
                CommandValue::None
            }
            WidgetCommand::GetValue => CommandValue::Number(self.position as f64),
            WidgetCommand::Custom(key, val) => {
                match key.as_str() {
                    "set_count" => {
                        if let CommandValue::Number(n) = val {
                            self.count = *n as i32;
                        }
                        CommandValue::None
                    }
                    "set_position_and_count" => {
                        if let CommandValue::Text(t) = val {
                            // "pos,count"
                            let parts: Vec<&str> = t.split(',').collect();
                            if parts.len() == 2 {
                                self.position = parts[0].parse().unwrap_or(0);
                                self.count = parts[1].parse().unwrap_or(0);
                            }
                        }
                        CommandValue::None
                    }
                    _ => CommandValue::None }
            }
            _ => CommandValue::None }
    }
}
