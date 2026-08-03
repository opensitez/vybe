//! MonthCalendar widget — calendar grid with header and day numbers.

use super::layout::{
    CommandValue, KeyEvent, LayoutRect, MouseButton as LayoutMouseButton, MouseEvent,
    MouseEventKind, PanelWidget, RenderContext, WidgetCommand, WidgetEvent, WidgetId };
use super::{WidgetColors, circle_path, rounded_rect_path};
use tiny_skia::*;

pub struct MonthCalendar {
    pub year: u32,
    pub month: u32,
    pub selected_day: u32,
    pub hover_day: Option<u32>,
    pub width: f32,
    pub height: f32,
    pub colors: WidgetColors,
    pub id: WidgetId,
    pub name: String,
    rect: LayoutRect,
    pending_events: Vec<WidgetEvent> }

impl MonthCalendar {
    pub fn new() -> Self {
        Self {
            year: 2026,
            month: 1,
            selected_day: 1,
            hover_day: None,
            width: 240.0,
            height: 200.0,
            colors: WidgetColors::default(),
            id: WidgetId::next(),
            name: String::new(),
            rect: LayoutRect::zero(),
            pending_events: Vec::new() }
    }

    pub fn with_name(mut self, name: &str) -> Self {
        self.name = name.to_string();
        self
    }

    /// Header height for month/year and nav arrows.
    fn header_height(&self) -> f32 {
        30.0
    }

    /// Day-of-week row height.
    fn dow_height(&self) -> f32 {
        20.0
    }

    /// Cell size for day grid.
    fn cell_size(&self) -> (f32, f32) {
        let w = self.width / 7.0;
        let grid_h = self.height - self.header_height() - self.dow_height();
        let h = grid_h / 6.0;
        (w, h)
    }

    /// Number of days in current month.
    fn days_in_month(&self) -> u32 {
        match self.month {
            1 | 3 | 5 | 7 | 8 | 10 | 12 => 31,
            4 | 6 | 9 | 11 => 30,
            2 => {
                if (self.year % 4 == 0 && self.year % 100 != 0) || self.year % 400 == 0 {
                    29
                } else {
                    28
                }
            }
            _ => 30 }
    }

    /// Day of week for the 1st of the month (0=Sunday).
    fn first_day_of_week(&self) -> u32 {
        // Zeller's congruence (simplified)
        let mut y = self.year as i32;
        let mut m = self.month as i32;
        if m < 3 {
            m += 12;
            y -= 1;
        }
        let q = 1;
        let k = y % 100;
        let j = y / 100;
        let h = (q + (13 * (m + 1)) / 5 + k + k / 4 + j / 4 - 2 * j) % 7;
        ((h + 6) % 7) as u32 // Convert to 0=Sunday
    }

    /// Paint — white background, header with nav arrows, day-of-week row, 6x7 day grid.
    /// All text (month name, day numbers, DOW labels) drawn by caller.
    pub fn paint(&self, pixmap: &mut Pixmap, x: f32, y: f32, scale: f32) {
        let ts = Transform::from_scale(scale, scale);
        let mut paint = Paint::default();
        paint.anti_alias = true;
        let mut stroke = Stroke::default();
        stroke.width = 1.0;
        let hh = self.header_height();
        let dh = self.dow_height();
        let (cw, ch) = self.cell_size();

        // White background
        paint.set_color_rgba8(255, 255, 255, 255);
        if let Some(path) = rounded_rect_path(x, y, self.width, self.height, 3.0) {
            pixmap.fill_path(&path, &paint, FillRule::Winding, ts, None);
        }

        // Header background
        let (ar, ag, ab, _) = self.colors.accent;
        paint.set_color_rgba8(ar, ag, ab, 255);
        if let Some(rect) = Rect::from_xywh(x + 1.0, y + 1.0, self.width - 2.0, hh - 1.0) {
            pixmap.fill_rect(rect, &paint, ts, None);
        }

        // Navigation arrows in header
        paint.set_color_rgba8(255, 255, 255, 255);
        let arrow_s = 5.0;
        let arrow_y = y + hh / 2.0;

        // Left arrow (<)
        let left_x = x + 12.0;
        let mut pb = PathBuilder::new();
        pb.move_to(left_x, arrow_y);
        pb.line_to(left_x + arrow_s, arrow_y - arrow_s);
        pb.line_to(left_x + arrow_s, arrow_y + arrow_s);
        pb.close();
        if let Some(path) = pb.finish() {
            pixmap.fill_path(&path, &paint, FillRule::Winding, ts, None);
        }

        // Right arrow (>)
        let right_x = x + self.width - 12.0;
        let mut pb = PathBuilder::new();
        pb.move_to(right_x, arrow_y);
        pb.line_to(right_x - arrow_s, arrow_y - arrow_s);
        pb.line_to(right_x - arrow_s, arrow_y + arrow_s);
        pb.close();
        if let Some(path) = pb.finish() {
            pixmap.fill_path(&path, &paint, FillRule::Winding, ts, None);
        }

        // Day-of-week row background
        paint.set_color_rgba8(245, 245, 245, 255);
        if let Some(rect) = Rect::from_xywh(x + 1.0, y + hh, self.width - 2.0, dh) {
            pixmap.fill_rect(rect, &paint, ts, None);
        }

        // Separator under DOW row
        paint.set_color_rgba8(220, 220, 220, 255);
        let mut pb = PathBuilder::new();
        pb.move_to(x, y + hh + dh);
        pb.line_to(x + self.width, y + hh + dh);
        if let Some(path) = pb.finish() {
            pixmap.stroke_path(&path, &paint, &stroke, ts, None);
        }

        // Day grid — highlight selected day and hover day
        let grid_y = y + hh + dh;
        let first_dow = self.first_day_of_week();
        let days = self.days_in_month();

        for day in 1..=days {
            let cell_idx = (first_dow + day - 1) as f32;
            let col = cell_idx % 7.0;
            let row = (cell_idx / 7.0).floor();
            let cx = x + col * cw + cw / 2.0;
            let cy = grid_y + row * ch + ch / 2.0;

            if day == self.selected_day {
                // Selected day: filled accent circle
                paint.set_color_rgba8(ar, ag, ab, 255);
                let r = (cw.min(ch) / 2.0 - 2.0).max(8.0);
                if let Some(path) = circle_path(cx, cy, r) {
                    pixmap.fill_path(&path, &paint, FillRule::Winding, ts, None);
                }
            } else if self.hover_day == Some(day) {
                // Hover day: light circle
                paint.set_color_rgba8(ar, ag, ab, 30);
                let r = (cw.min(ch) / 2.0 - 2.0).max(8.0);
                if let Some(path) = circle_path(cx, cy, r) {
                    pixmap.fill_path(&path, &paint, FillRule::Winding, ts, None);
                }
            }
        }

        // Outer border
        paint.set_color_rgba8(180, 180, 180, 255);
        if let Some(path) = rounded_rect_path(x, y, self.width, self.height, 3.0) {
            pixmap.stroke_path(&path, &paint, &stroke, ts, None);
        }
    }

    pub fn measure(&self) -> (f32, f32) {
        (self.width, self.height)
    }

    /// Hit test — returns day number at position, or None.
    pub fn hit_test(&self, mx: f32, my: f32) -> Option<u32> {
        let hh = self.header_height();
        let dh = self.dow_height();
        let (cw, ch) = self.cell_size();
        let grid_y = hh + dh;

        if my < grid_y || my > self.height {
            return None;
        }
        if mx < 0.0 || mx > self.width {
            return None;
        }

        let col = (mx / cw) as u32;
        let row = ((my - grid_y) / ch) as u32;
        if col >= 7 || row >= 6 {
            return None;
        }

        let cell_idx = row * 7 + col;
        let first_dow = self.first_day_of_week();
        if cell_idx < first_dow {
            return None;
        }
        let day = cell_idx - first_dow + 1;
        if day >= 1 && day <= self.days_in_month() {
            Some(day)
        } else {
            None
        }
    }

    /// Cell position for a given day (for text placement by caller).
    pub fn day_cell(&self, day: u32) -> (f32, f32, f32, f32) {
        let hh = self.header_height();
        let dh = self.dow_height();
        let (cw, ch) = self.cell_size();
        let first_dow = self.first_day_of_week();
        let cell_idx = first_dow + day - 1;
        let col = cell_idx % 7;
        let row = cell_idx / 7;
        (col as f32 * cw, hh + dh + row as f32 * ch, cw, ch)
    }

    /// Navigate to previous month.
    pub fn prev_month(&mut self) {
        if self.month == 1 {
            self.month = 12;
            self.year -= 1;
        } else {
            self.month -= 1;
        }
        self.selected_day = self.selected_day.min(self.days_in_month());
    }

    /// Navigate to next month.
    pub fn next_month(&mut self) {
        if self.month == 12 {
            self.month = 1;
            self.year += 1;
        } else {
            self.month += 1;
        }
        self.selected_day = self.selected_day.min(self.days_in_month());
    }
}

impl PanelWidget for MonthCalendar {
    fn name(&self) -> &str {
        &self.name
    }
    fn widget_id(&self) -> WidgetId {
        self.id
    }
    fn set_rect(&mut self, rect: LayoutRect) {
        self.rect = rect;
        self.width = rect.w;
        self.height = rect.h;
    }
    fn rect(&self) -> LayoutRect {
        self.rect
    }

    fn render(&mut self, ctx: &mut RenderContext) {
        let r = self.rect;
        if r.w <= 0.0 || r.h <= 0.0 {
            return;
        }
        self.paint(ctx.pixmap, r.x, r.y, ctx.scale);
        // Draw month/year header text
        let months = [
            "January",
            "February",
            "March",
            "April",
            "May",
            "June",
            "July",
            "August",
            "September",
            "October",
            "November",
            "December",
        ];
        let header = format!(
            "{} {}",
            months
                .get(self.month.saturating_sub(1) as usize)
                .unwrap_or(&"?"),
            self.year
        );
        let white = cosmic_text::Color::rgba(255, 255, 255, 255);
        super::ide_text::draw_text(
            ctx.pixmap,
            ctx.font_system,
            ctx.swash_cache,
            &header,
            r.x + r.w / 2.0 - 40.0,
            r.y + 6.0,
            13.0,
            white,
            ctx.scale,
        );
        // Day-of-week labels
        let dow = ["Su", "Mo", "Tu", "We", "Th", "Fr", "Sa"];
        let (cw, _ch) = self.cell_size();
        let gray = cosmic_text::Color::rgba(80, 80, 80, 255);
        for (i, d) in dow.iter().enumerate() {
            super::ide_text::draw_text(
                ctx.pixmap,
                ctx.font_system,
                ctx.swash_cache,
                d,
                r.x + i as f32 * cw + cw / 2.0 - 6.0,
                r.y + self.header_height() + 2.0,
                11.0,
                gray,
                ctx.scale,
            );
        }
        // Day numbers
        let dark = cosmic_text::Color::rgba(30, 30, 30, 255);
        for day in 1..=self.days_in_month() {
            let (cx, cy, dw, dh) = self.day_cell(day);
            let col = if day == self.selected_day {
                white
            } else {
                dark
            };
            let txt = format!("{}", day);
            super::ide_text::draw_text(
                ctx.pixmap,
                ctx.font_system,
                ctx.swash_cache,
                &txt,
                r.x + cx + dw / 2.0 - 5.0,
                r.y + cy + dh / 2.0 - 6.0,
                12.0,
                col,
                ctx.scale,
            );
        }
    }

    fn handle_mouse(&mut self, event: &MouseEvent) -> bool {
        let r = self.rect;
        if !r.contains(event.x, event.y) {
            return false;
        }
        let lx = event.x - r.x;
        let ly = event.y - r.y;
        match event.kind {
            MouseEventKind::Press(LayoutMouseButton::Left) => {
                // Check nav arrows
                if ly < self.header_height() {
                    if lx < 24.0 {
                        self.prev_month();
                        return true;
                    }
                    if lx > self.width - 24.0 {
                        self.next_month();
                        return true;
                    }
                    return false;
                }
                if let Some(day) = self.hit_test(lx, ly) {
                    self.selected_day = day;
                    self.pending_events
                        .push(WidgetEvent::CalendarDateSelected(self.name.clone(), day));
                    return true;
                }
            }
            MouseEventKind::Move => {
                self.hover_day = self.hit_test(lx, ly);
            }
            _ => {}
        }
        false
    }

    fn handle_key(&mut self, _event: &KeyEvent) -> bool {
        false
    }
    fn focusable(&self) -> bool {
        true
    }
    fn handle_command(&mut self, cmd: &WidgetCommand) -> CommandValue {
        match cmd {
            WidgetCommand::SetValue(v) => {
                self.selected_day = *v as u32;
                CommandValue::None
            }
            WidgetCommand::GetValue => CommandValue::Number(self.selected_day as f64),
            _ => CommandValue::None }
    }

    fn drain_events(&mut self) -> Vec<WidgetEvent> {
        std::mem::take(&mut self.pending_events)
    }
}
