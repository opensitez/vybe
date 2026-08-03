//! OutputPanel — a bottom panel with Output and Problems tabs.
//!
//! Renders a tab bar (Output | Problems), content area with scrollable lines,
//! and close/clear buttons. Emits events for tab switching, close, and item clicks.

use crate::ide_text;
use crate::layout::*;
use cosmic_text::Color as CosmicColor;
use tiny_skia::*;

/// Diagnostic severity for the Problems panel.
#[derive(Clone, Copy, Debug, PartialEq)]
pub enum ProblemSeverity {
    Error,
    Warning,
    Info,
    Hint }

/// A single problem entry.
#[derive(Clone, Debug)]
pub struct ProblemEntry {
    pub file: String,
    pub line: usize,
    pub severity: ProblemSeverity,
    pub message: String }

/// Which tab is active in the output panel.
#[derive(Clone, Copy, PartialEq, Debug)]
pub enum OutputTab {
    Output,
    Problems }

/// Events emitted by the OutputPanel.
#[derive(Clone, Debug)]
pub enum OutputPanelEvent {
    /// User clicked close button
    Close,
    /// User clicked clear button (Output tab)
    ClearOutput,
    /// User switched tab
    TabChanged(OutputTab),
    /// User clicked a problem at the given index
    ProblemClicked(usize) }

pub struct OutputPanel {
    pub id: WidgetId,
    rect: LayoutRect,
    active_tab: OutputTab,
    output_lines: Vec<String>,
    problems: Vec<ProblemEntry>,
    scroll_y: f32,
    pending_events: Vec<WidgetEvent>,
    bg_color: (u8, u8, u8, u8),
    header_bg: (u8, u8, u8, u8),
    accent_color: (u8, u8, u8, u8),
    visible: bool }

impl OutputPanel {
    pub fn new() -> Self {
        Self {
            id: WidgetId::next(),
            rect: LayoutRect::zero(),
            active_tab: OutputTab::Output,
            output_lines: Vec::new(),
            problems: Vec::new(),
            scroll_y: 0.0,
            pending_events: Vec::new(),
            bg_color: (25, 25, 30, 255),
            header_bg: (35, 35, 42, 255),
            accent_color: (0, 122, 204, 255),
            visible: true }
    }

    pub fn set_output_lines(&mut self, lines: &[String]) {
        self.output_lines.clear();
        self.output_lines.extend(lines.iter().cloned());
    }

    pub fn set_problems(&mut self, problems: Vec<ProblemEntry>) {
        self.problems = problems;
    }

    pub fn set_active_tab(&mut self, tab: OutputTab) {
        if tab != self.active_tab {
            self.active_tab = tab;
            self.scroll_y = 0.0;
        }
    }

    pub fn active_tab(&self) -> OutputTab {
        self.active_tab
    }

    pub fn set_visible(&mut self, v: bool) {
        self.visible = v;
    }
    pub fn visible(&self) -> bool {
        self.visible
    }

    pub fn clear_output(&mut self) {
        self.output_lines.clear();
        self.scroll_y = 0.0;
    }

    pub fn scroll_y(&self) -> f32 {
        self.scroll_y
    }
    pub fn set_scroll_y(&mut self, y: f32) {
        self.scroll_y = y;
    }

    /// Scroll to the last line of the active tab's content.
    pub fn scroll_to_bottom(&mut self) {
        let total = match self.active_tab {
            OutputTab::Output => self.output_lines.len(),
            OutputTab::Problems => self.problems.len() };
        self.scroll_y = (total as f32 * Self::LINE_H - (self.rect.h - Self::HEADER_H)).max(0.0);
    }

    /// Drain pending OutputPanelEvent actions from WidgetEvent::Action
    pub fn drain_panel_events(&mut self) -> Vec<OutputPanelEvent> {
        let mut result = Vec::new();
        let events: Vec<WidgetEvent> = self.pending_events.drain(..).collect();
        for evt in events {
            if let WidgetEvent::Action(s) = evt {
                match s.as_str() {
                    "close" => result.push(OutputPanelEvent::Close),
                    "clear" => result.push(OutputPanelEvent::ClearOutput),
                    "tab_output" => result.push(OutputPanelEvent::TabChanged(OutputTab::Output)),
                    "tab_problems" => {
                        result.push(OutputPanelEvent::TabChanged(OutputTab::Problems))
                    }
                    _ if s.starts_with("problem_") => {
                        if let Ok(idx) = s[8..].parse::<usize>() {
                            result.push(OutputPanelEvent::ProblemClicked(idx));
                        }
                    }
                    _ => {}
                }
            }
        }
        result
    }

    const HEADER_H: f32 = 24.0;
    const LINE_H: f32 = 18.0;
}

impl PanelWidget for OutputPanel {
    fn set_rect(&mut self, rect: LayoutRect) {
        self.rect = rect;
    }
    fn rect(&self) -> LayoutRect {
        self.rect
    }
    fn widget_id(&self) -> WidgetId {
        self.id
    }

    fn render(&mut self, ctx: &mut RenderContext) {
        if !self.visible {
            return;
        }
        let s = ctx.scale;
        let mut paint = Paint::default();

        // Background
        let (r, g, b, a) = self.bg_color;
        paint.set_color_rgba8(r, g, b, a);
        if let Some(rect) = Rect::from_xywh(
            self.rect.x * s,
            self.rect.y * s,
            self.rect.w * s,
            self.rect.h * s,
        ) {
            ctx.pixmap
                .fill_rect(rect, &paint, Transform::identity(), None);
        }

        // Header bar
        let (r, g, b, a) = self.header_bg;
        paint.set_color_rgba8(r, g, b, a);
        if let Some(rect) = Rect::from_xywh(
            self.rect.x * s,
            self.rect.y * s,
            self.rect.w * s,
            Self::HEADER_H * s,
        ) {
            ctx.pixmap
                .fill_rect(rect, &paint, Transform::identity(), None);
        }

        // Tab buttons: Output | Problems
        let tab_y = self.rect.y + 4.0;
        let output_tab_x = self.rect.x + 10.0;
        let output_col = if self.active_tab == OutputTab::Output {
            CosmicColor::rgba(230, 230, 230, 255)
        } else {
            CosmicColor::rgba(120, 120, 120, 255)
        };
        ide_text::draw_text(
            ctx.pixmap,
            ctx.font_system,
            ctx.swash_cache,
            "Output",
            output_tab_x,
            tab_y,
            13.0,
            output_col,
            s,
        );

        // Active underline for Output
        if self.active_tab == OutputTab::Output {
            let (ar, ag, ab, _) = self.accent_color;
            paint.set_color_rgba8(ar, ag, ab, 255);
            if let Some(rect) = Rect::from_xywh(
                output_tab_x * s,
                (self.rect.y + 22.0) * s,
                50.0 * s,
                2.0 * s,
            ) {
                ctx.pixmap
                    .fill_rect(rect, &paint, Transform::identity(), None);
            }
        }

        let problems_tab_x = self.rect.x + 80.0;
        let problems_label = format!("Problems ({})", self.problems.len());
        let problems_col = if self.active_tab == OutputTab::Problems {
            CosmicColor::rgba(230, 230, 230, 255)
        } else {
            CosmicColor::rgba(120, 120, 120, 255)
        };
        ide_text::draw_text(
            ctx.pixmap,
            ctx.font_system,
            ctx.swash_cache,
            &problems_label,
            problems_tab_x,
            tab_y,
            13.0,
            problems_col,
            s,
        );

        if self.active_tab == OutputTab::Problems {
            let (ar, ag, ab, _) = self.accent_color;
            paint.set_color_rgba8(ar, ag, ab, 255);
            if let Some(rect) = Rect::from_xywh(
                problems_tab_x * s,
                (self.rect.y + 22.0) * s,
                100.0 * s,
                2.0 * s,
            ) {
                ctx.pixmap
                    .fill_rect(rect, &paint, Transform::identity(), None);
            }
        }

        // Close button (×)
        let close_x = self.rect.x + self.rect.w - 24.0;
        ide_text::draw_text(
            ctx.pixmap,
            ctx.font_system,
            ctx.swash_cache,
            "×",
            close_x,
            tab_y,
            13.0,
            CosmicColor::rgba(150, 150, 150, 255),
            s,
        );

        // Clear button (only for Output tab)
        if self.active_tab == OutputTab::Output {
            let clear_x = self.rect.x + self.rect.w - 80.0;
            ide_text::draw_text(
                ctx.pixmap,
                ctx.font_system,
                ctx.swash_cache,
                "Clear",
                clear_x,
                tab_y,
                13.0,
                CosmicColor::rgba(120, 120, 120, 255),
                s,
            );
        }

        // Separator line at top
        let mut sep = Paint::default();
        sep.set_color_rgba8(60, 60, 70, 255);
        if let Some(rect) = Rect::from_xywh(self.rect.x * s, self.rect.y * s, self.rect.w * s, 1.0)
        {
            ctx.pixmap
                .fill_rect(rect, &sep, Transform::identity(), None);
        }

        // Content area
        let content_y = self.rect.y + Self::HEADER_H;
        let content_h = self.rect.h - Self::HEADER_H;
        let visible_lines = (content_h / Self::LINE_H) as usize;

        match self.active_tab {
            OutputTab::Output => {
                let skip = (self.scroll_y / Self::LINE_H).max(0.0) as usize;
                for (i, line) in self
                    .output_lines
                    .iter()
                    .skip(skip)
                    .take(visible_lines + 1)
                    .enumerate()
                {
                    let ly = content_y + i as f32 * Self::LINE_H - (self.scroll_y % Self::LINE_H);
                    if ly >= content_y && ly < self.rect.y + self.rect.h {
                        let col = if line.starts_with("ERR:") || line.starts_with("Save error") {
                            CosmicColor::rgba(255, 100, 100, 255)
                        } else if line.starts_with("Building") || line.starts_with("Running") {
                            CosmicColor::rgba(100, 200, 100, 255)
                        } else {
                            CosmicColor::rgba(180, 180, 180, 255)
                        };
                        let display = if line.len() > 120 {
                            &line[..120]
                        } else {
                            line.as_str()
                        };
                        ide_text::draw_text(
                            ctx.pixmap,
                            ctx.font_system,
                            ctx.swash_cache,
                            display,
                            self.rect.x + 10.0,
                            ly + 2.0,
                            13.0,
                            col,
                            s,
                        );
                    }
                }
            }
            OutputTab::Problems => {
                if self.problems.is_empty() {
                    ide_text::draw_text(
                        ctx.pixmap,
                        ctx.font_system,
                        ctx.swash_cache,
                        "No problems detected.",
                        self.rect.x + 10.0,
                        content_y + 4.0,
                        13.0,
                        CosmicColor::rgba(100, 200, 100, 255),
                        s,
                    );
                } else {
                    let skip = (self.scroll_y / Self::LINE_H).max(0.0) as usize;
                    for (i, prob) in self
                        .problems
                        .iter()
                        .skip(skip)
                        .take(visible_lines + 1)
                        .enumerate()
                    {
                        let ly =
                            content_y + i as f32 * Self::LINE_H - (self.scroll_y % Self::LINE_H);
                        if ly >= content_y && ly < self.rect.y + self.rect.h {
                            let (icon, icon_col) = match prob.severity {
                                ProblemSeverity::Error => {
                                    ("●", CosmicColor::rgba(255, 80, 80, 255))
                                }
                                ProblemSeverity::Warning => {
                                    ("▲", CosmicColor::rgba(255, 200, 50, 255))
                                }
                                ProblemSeverity::Info => {
                                    ("ℹ", CosmicColor::rgba(80, 160, 255, 255))
                                }
                                ProblemSeverity::Hint => {
                                    ("…", CosmicColor::rgba(140, 140, 140, 255))
                                }
                            };
                            // Icon
                            ide_text::draw_text(
                                ctx.pixmap,
                                ctx.font_system,
                                ctx.swash_cache,
                                icon,
                                self.rect.x + 10.0,
                                ly + 2.0,
                                13.0,
                                icon_col,
                                s,
                            );
                            // File:line
                            let loc = format!("{}:{}", prob.file, prob.line);
                            ide_text::draw_text(
                                ctx.pixmap,
                                ctx.font_system,
                                ctx.swash_cache,
                                &loc,
                                self.rect.x + 24.0,
                                ly + 2.0,
                                13.0,
                                CosmicColor::rgba(130, 180, 230, 255),
                                s,
                            );
                            // Message
                            let msg = if prob.message.len() > 100 {
                                &prob.message[..100]
                            } else {
                                prob.message.as_str()
                            };
                            ide_text::draw_text(
                                ctx.pixmap,
                                ctx.font_system,
                                ctx.swash_cache,
                                msg,
                                self.rect.x + 200.0,
                                ly + 2.0,
                                13.0,
                                CosmicColor::rgba(200, 200, 200, 255),
                                s,
                            );
                        }
                    }
                }
            }
        }
    }

    fn handle_mouse(&mut self, event: &MouseEvent) -> bool {
        if !self.visible {
            return false;
        }
        if !self.rect.contains(event.x, event.y) {
            return false;
        }

        match event.kind {
            MouseEventKind::Press(MouseButton::Left) => {
                let rel_y = event.y - self.rect.y;

                // Header area
                if rel_y < Self::HEADER_H {
                    let rel_x = event.x - self.rect.x;

                    // Close button
                    if rel_x > self.rect.w - 24.0 {
                        self.pending_events
                            .push(WidgetEvent::Action("close".into()));
                        return true;
                    }

                    // Clear button (Output tab only)
                    if self.active_tab == OutputTab::Output
                        && rel_x > self.rect.w - 80.0
                        && rel_x < self.rect.w - 30.0
                    {
                        self.pending_events
                            .push(WidgetEvent::Action("clear".into()));
                        return true;
                    }

                    // Tab switching
                    if rel_x >= 10.0 && rel_x < 70.0 {
                        self.active_tab = OutputTab::Output;
                        self.scroll_y = 0.0;
                        self.pending_events
                            .push(WidgetEvent::Action("tab_output".into()));
                        return true;
                    }
                    if rel_x >= 80.0 && rel_x < 200.0 {
                        self.active_tab = OutputTab::Problems;
                        self.scroll_y = 0.0;
                        self.pending_events
                            .push(WidgetEvent::Action("tab_problems".into()));
                        return true;
                    }
                    return true;
                }

                // Content area click (problems navigation)
                if self.active_tab == OutputTab::Problems && rel_y >= Self::HEADER_H {
                    let content_rel_y = rel_y - Self::HEADER_H;
                    let clicked_idx = (self.scroll_y / Self::LINE_H).max(0.0) as usize
                        + ((content_rel_y + (self.scroll_y % Self::LINE_H)) / Self::LINE_H)
                            as usize;
                    if clicked_idx < self.problems.len() {
                        self.pending_events
                            .push(WidgetEvent::Action(format!("problem_{}", clicked_idx)));
                    }
                    return true;
                }

                true
            }
            _ => true, // absorb all mouse events in our area
        }
    }

    fn handle_key(&mut self, _event: &KeyEvent) -> bool {
        false
    }

    fn handle_scroll(&mut self, delta: f32, _x: f32, _y: f32) -> bool {
        if !self.visible {
            return false;
        }
        self.scroll_y = (self.scroll_y - delta).max(0.0);
        let total_lines = match self.active_tab {
            OutputTab::Output => self.output_lines.len(),
            OutputTab::Problems => self.problems.len() };
        let max_scroll =
            (total_lines as f32 * Self::LINE_H - (self.rect.h - Self::HEADER_H)).max(0.0);
        self.scroll_y = self.scroll_y.min(max_scroll);
        true
    }

    fn cursor_at(&self, _x: f32, _y: f32) -> winit::window::CursorIcon {
        winit::window::CursorIcon::Default
    }

    fn drain_events(&mut self) -> Vec<WidgetEvent> {
        self.pending_events.drain(..).collect()
    }
}
