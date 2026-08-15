//! Label widget — standalone tiny-skia rendered label.

use super::WidgetColors;
use super::layout::{
    CommandValue, KeyEvent, LayoutRect, MouseEvent, PanelWidget, RenderContext, TextAlign,
    WidgetCommand, WidgetId, command_color, command_number,
};
use super::ide_text::FontSpec;
use cosmic_text::Color as CosmicColor;
use tiny_skia::*;

pub struct Label {
    pub id: WidgetId,
    pub name: String,
    pub text: String,
    pub width: f32,
    pub height: f32,
    pub auto_size: bool,
    pub transparent: bool,
    pub colors: WidgetColors,
    /// The label's resolved text style — family, size, weight, slant and
    /// decoration together, because they are shaped together. A size alone
    /// could not spell `<b>`, `<i>`, `<code>` or `<a>`.
    pub font: FontSpec,
    pub text_align: TextAlign,
    pub word_wrap: bool,
    rect: LayoutRect,
}

impl Label {
    pub fn new<S: Into<String>>(text: S) -> Self {
        Self {
            id: WidgetId::next(),
            name: String::new(),
            text: text.into(),
            width: 100.0,
            height: 20.0,
            auto_size: true,
            transparent: true,
            colors: WidgetColors {
                background: (240, 240, 240, 255),
                ..WidgetColors::default()
            },
            font: FontSpec::sans(13.0),
            text_align: TextAlign::Left,
            word_wrap: false,
            rect: LayoutRect::zero(),
        }
    }

    pub fn with_name<S: Into<String>>(mut self, name: S) -> Self {
        self.name = name.into();
        self
    }
    pub fn with_text_align(mut self, align: TextAlign) -> Self {
        self.text_align = align;
        self
    }
    pub fn with_word_wrap(mut self) -> Self {
        self.word_wrap = true;
        self
    }

    /// Paint the label background (if not transparent). Text drawn by caller.
    pub fn paint(&self, pixmap: &mut Pixmap, x: f32, y: f32, scale: f32) {
        if self.transparent {
            return;
        }
        let ts = Transform::from_scale(scale, scale);
        let mut paint = Paint::default();
        paint.anti_alias = true;

        // Background fill
        let (r, g, b, a) = self.colors.background;
        paint.set_color_rgba8(r, g, b, a);
        if let Some(rect) = Rect::from_xywh(x, y, self.width, self.height) {
            pixmap.fill_rect(rect, &paint, ts, None);
        }
    }

    pub fn measure(&self) -> (f32, f32) {
        (self.width, self.height)
    }

    /// Compute the x position for text given alignment.
    fn align_x(
        &self,
        font_system: &mut cosmic_text::FontSystem,
        text: &str,
        left: f32,
        available_w: f32,
        scale: f32,
    ) -> f32 {
        match self.text_align {
            TextAlign::Left => left,
            TextAlign::Center => {
                let tw = super::ide_text::measure_text_spec(font_system, text, &self.font, scale);
                left + (available_w - tw).max(0.0) / 2.0
            }
            TextAlign::Right => {
                let tw = super::ide_text::measure_text_spec(font_system, text, &self.font, scale);
                left + (available_w - tw).max(0.0)
            }
        }
    }

    /// Word-wrap text into lines fitting within `max_width`.
    fn wrap_text(
        &self,
        font_system: &mut cosmic_text::FontSystem,
        max_width: f32,
        scale: f32,
    ) -> Vec<String> {
        let mut lines = Vec::new();
        for paragraph in self.text.split('\n') {
            let words: Vec<&str> = paragraph.split_whitespace().collect();
            if words.is_empty() {
                lines.push(String::new());
                continue;
            }
            let mut current_line = String::new();
            for word in words {
                if current_line.is_empty() {
                    current_line = word.to_string();
                } else {
                    let candidate = format!("{} {}", current_line, word);
                    let w = super::ide_text::measure_text_spec(
                        font_system,
                        &candidate,
                        &self.font,
                        scale,
                    );
                    if w <= max_width {
                        current_line = candidate;
                    } else {
                        lines.push(current_line);
                        current_line = word.to_string();
                    }
                }
            }
            lines.push(current_line);
        }
        lines
    }
}

// ── PanelWidget impl ───────────────────────────────────────────────────

impl PanelWidget for Label {
    fn set_rect(&mut self, rect: LayoutRect) {
        self.rect = rect;
        self.width = rect.w;
        self.height = rect.h;
    }

    fn rect(&self) -> LayoutRect {
        self.rect
    }
    fn widget_id(&self) -> WidgetId {
        self.id
    }
    fn name(&self) -> &str {
        &self.name
    }

    fn render(&mut self, ctx: &mut RenderContext) {
        let r = self.rect;
        if r.w <= 0.0 || r.h <= 0.0 {
            return;
        }

        // Background (if not transparent)
        self.paint(ctx.pixmap, r.x, r.y, ctx.scale);

        let (fr, fg, fb, _) = self.colors.foreground;
        let color = CosmicColor::rgba(fr, fg, fb, 255);
        let padding = 2.0;
        let available_w = r.w - padding * 2.0;

        if self.word_wrap {
            // Word-wrap: split text into lines that fit within available_w
            let lines = self.wrap_text(ctx.font_system, available_w, ctx.scale);
            let mut ty = r.y + padding;
            let line_height = self.font.size + 2.0;
            for line in &lines {
                if ty + line_height > r.y + r.h {
                    break;
                }
                let tx = self.align_x(ctx.font_system, line, r.x + padding, available_w, ctx.scale);
                super::ide_text::draw_text_spec(
                    ctx.pixmap,
                    ctx.font_system,
                    ctx.swash_cache,
                    line,
                    tx,
                    ty,
                    &self.font,
                    color,
                    ctx.scale,
                );
                ty += line_height;
            }
        } else {
            let ty = r.y + (r.h - self.font.size) / 2.0 - 1.0;
            let tx = self.align_x(
                ctx.font_system,
                &self.text,
                r.x + padding,
                available_w,
                ctx.scale,
            );
            super::ide_text::draw_text_spec(
                ctx.pixmap,
                ctx.font_system,
                ctx.swash_cache,
                &self.text,
                tx,
                ty,
                &self.font,
                color,
                ctx.scale,
            );
        }
    }

    fn handle_mouse(&mut self, _event: &MouseEvent) -> bool {
        false
    }
    fn handle_key(&mut self, _event: &KeyEvent) -> bool {
        false
    }

    fn handle_command(&mut self, cmd: &WidgetCommand) -> CommandValue {
        match cmd {
            WidgetCommand::SetText(t) => {
                self.text = t.clone();
                CommandValue::None
            }
            WidgetCommand::GetText => CommandValue::Text(self.text.clone()),
            // Text styling and explicit sizing — a `Text(style: TextStyle(…))`
            // adapter drives colour/size here, and a rule-like label (a
            // `Divider`) gets its thickness from an explicit height.
            WidgetCommand::Custom(key, val) => match key.as_str() {
                "SetForeColor" => {
                    if let Some(rgba) = command_color(val) {
                        self.colors.foreground = rgba;
                    }
                    CommandValue::None
                }
                "SetBackColor" => {
                    if let Some(rgba) = command_color(val) {
                        self.colors.background = rgba;
                        // An explicit background is only visible if the label
                        // stops painting itself as transparent.
                        self.transparent = false;
                    }
                    CommandValue::None
                }
                // The whole font axis, in one line — see `FontSpec::apply_command`.
                key if self.font.apply_command(key, val) => CommandValue::None,
                "SetTextAlign" => {
                    if let CommandValue::Text(value) = val {
                        if let Some(align) = TextAlign::from_css(value) {
                            self.text_align = align;
                        }
                    }
                    CommandValue::None
                }
                "SetWidth" => {
                    if let Some(w) = command_number(val) {
                        self.width = w as f32;
                        self.auto_size = false;
                    }
                    CommandValue::None
                }
                "SetHeight" => {
                    if let Some(h) = command_number(val) {
                        self.height = h as f32;
                        self.auto_size = false;
                    }
                    CommandValue::None
                }
                _ => CommandValue::None,
            },
            _ => CommandValue::None,
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn styled(command: &str, value: &str) -> Label {
        let mut label = Label::new("text");
        label.handle_command(&WidgetCommand::Custom(
            command.into(),
            CommandValue::Text(value.into()),
        ));
        label
    }

    #[test]
    fn bold_is_a_weight_whether_it_is_spelled_as_a_word_or_a_number() {
        // CSS `font-weight` is a number, and `bold` is a NAME for 700 — so both
        // spellings are the same property arriving by the same route. Treating
        // the keyword as a separate boolean is how a toolkit ends up unable to
        // express `font-weight: 600`.
        assert_eq!(styled("SetFontWeight", "bold").font.weight, 700);
        assert_eq!(styled("SetFontWeight", "600").font.weight, 600);
        assert_eq!(styled("SetFontWeight", "normal").font.weight, 400);
    }

    #[test]
    fn oblique_slants_the_same_as_italic() {
        assert!(styled("SetFontStyle", "italic").font.italic);
        assert!(styled("SetFontStyle", "oblique").font.italic);
        assert!(!styled("SetFontStyle", "normal").font.italic);
    }

    #[test]
    fn text_decoration_is_a_list_so_both_lines_can_be_asked_for_at_once() {
        let both = styled("SetTextDecoration", "underline line-through");
        assert!(both.font.underline && both.font.line_through);

        let one = styled("SetTextDecoration", "underline");
        assert!(one.font.underline && !one.font.line_through);

        let neither = styled("SetTextDecoration", "none");
        assert!(!neither.font.underline && !neither.font.line_through);
    }

    #[test]
    fn a_family_reaches_the_label_unresolved() {
        // Stored as the CSS name, not a resolved font: which face `monospace`
        // means is a question for the font database at draw time, and the
        // declaration has to survive being read back.
        assert_eq!(styled("SetFontFamily", "monospace").font.family, "monospace");
    }
}
