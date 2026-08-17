//! Progress bar widget

use super::layout::{
    CommandValue, KeyEvent, LayoutRect, MouseEvent, PanelWidget, RenderContext, WidgetCommand,
    WidgetId,
};
use super::{WidgetColors, rounded_rect_path};
use tiny_skia::*;

pub struct ProgressBar {
    /// `<progress value>` — in the units `max` sets, NOT a fraction.
    ///
    /// It used to be documented as `0.0..1.0` and clamped there on the way in,
    /// which silently destroyed the value: `<progress value="60" max="100">` —
    /// the ordinary spelling — arrived as `60`, saturated to `1.0`, and drew a
    /// FULL bar for what should have been 60%.
    pub value: f32,
    /// `<progress max>` — HTML §4.10.13. **Defaults to 1**, which is what makes
    /// a bare `value="0.6"` mean 60% and kept the old clamp looking correct for
    /// the one case anybody tested.
    pub max: f32,
    pub width: f32,
    pub height: f32,
    pub colors: WidgetColors,
    pub id: WidgetId,
    pub name: String,
    rect: LayoutRect,
}

impl ProgressBar {
    pub fn new() -> Self {
        Self {
            value: 0.0,
            max: 1.0,
            width: 200.0,
            height: 16.0,
            colors: WidgetColors::default(),
            id: WidgetId::next(),
            name: String::new(),
            rect: LayoutRect::zero(),
        }
    }

    pub fn with_name(mut self, name: &str) -> Self {
        self.name = name.to_string();
        self
    }

    /// How full the bar is, `0.0..=1.0` — `value / max`, HTML §4.10.13.
    ///
    /// A `max` of zero or less is not a division to guard against elsewhere:
    /// the spec says a non-positive `max` makes the element indeterminate, and
    /// an empty bar is the honest drawing of "no answer".
    fn fraction(&self) -> f32 {
        if self.max <= 0.0 {
            return 0.0;
        }
        (self.value / self.max).clamp(0.0, 1.0)
    }

    pub fn paint(&self, pixmap: &mut Pixmap, x: f32, y: f32, scale: f32) {
        let ts = Transform::from_scale(scale, scale);
        let mut paint = Paint::default();
        paint.anti_alias = true;

        // Background
        let (r, g, b, a) = (240, 240, 240, 255);
        paint.set_color_rgba8(r, g, b, a);
        if let Some(path) = rounded_rect_path(x, y, self.width, self.height, 4.0) {
            pixmap.fill_path(&path, &paint, FillRule::Winding, ts, None);
        }

        // Filled portion — the FRACTION is `value / max`, which is what makes
        // `value="60" max="100"` six tenths of the bar rather than all of it.
        let fill_w = self.fraction() * self.width;
        let (fr, fg, fb, fa) = self.colors.accent;
        paint.set_color_rgba8(fr, fg, fb, fa);
        if fill_w > 0.0 {
            if let Some(path) = rounded_rect_path(x, y, fill_w, self.height, 4.0) {
                pixmap.fill_path(&path, &paint, FillRule::Winding, ts, None);
            }
        }

        // Border
        let (br, bg, bb, ba) = self.colors.border;
        paint.set_color_rgba8(br, bg, bb, ba);
        let mut stroke = Stroke::default();
        stroke.width = 1.0;
        if let Some(path) = rounded_rect_path(x, y, self.width, self.height, 4.0) {
            pixmap.stroke_path(&path, &paint, &stroke, ts, None);
        }
    }

    pub fn measure(&self) -> (f32, f32) {
        (self.width, self.height)
    }
}

impl PanelWidget for ProgressBar {
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
    }
    fn handle_mouse(&mut self, _event: &MouseEvent) -> bool {
        false
    }
    fn handle_key(&mut self, _event: &KeyEvent) -> bool {
        false
    }

    fn handle_command(&mut self, cmd: &WidgetCommand) -> CommandValue {
        match cmd {
            WidgetCommand::SetValue(v) => {
                // **Stored as written, clamped only when drawn.** Clamping on
                // the way IN makes the answer depend on attribute ORDER, which
                // no author controls: `value="60"` then `max="100"` clamped 60
                // to the default max of 1, and the later `max` could not undo
                // it because the 60 was already gone. Range is a question about
                // the fraction, so `fraction()` is where it is asked.
                self.value = *v as f32;
                CommandValue::None
            }
            WidgetCommand::GetValue => CommandValue::Number(self.value as f64),
            // `<progress max>`. There was no arm at all, so `max="100"` was
            // dropped and the default of 1 stood — the reason a 60-of-100 bar
            // came out full.
            WidgetCommand::Custom(name, value) if name == "SetMax" => {
                if let Some(max) = crate::layout::command_number(value) {
                    self.max = max as f32;
                }
                CommandValue::None
            }
            _ => CommandValue::None,
        }
    }
}
