//! `<input type="file">` — the file picker control.
//!
//! HTML renders this as a button plus the chosen file's name, and clicking the
//! button opens the platform's file chooser. That is also what VB's
//! `OpenFileDialog` is, which is why the control belongs here rather than in a
//! frontend: one control, two spellings.
//!
//! Before this, `<input type="file">` fell to `control_kind`'s text arm and
//! rendered an editable text box — a field you could type prose into, which
//! neither picks a file nor reports one. The native chooser was already
//! available ([`crate::dialogs::FileDialog`], over `rfd`); nothing reached it
//! from an element.

use super::layout::{
    CommandValue, KeyEvent, LayoutRect, MouseButton as LayoutMouseButton, MouseEvent,
    MouseEventKind, PanelWidget, RenderContext, WidgetCommand, WidgetEvent, WidgetId,
};
use super::{WidgetColors, rounded_rect_path};
use tiny_skia::*;

/// The label on the button half, matching what a browser shows.
const CHOOSE: &str = "Choose File";
/// What the text half says with nothing picked — HTML's own wording.
const NOTHING: &str = "No file chosen";
/// How wide the button half is. Fixed, because the text half takes whatever is
/// left: a filename is the part that varies.
const BUTTON_W: f32 = 96.0;

pub struct FileInput {
    /// The chosen path, or empty. This IS the control's `value`.
    ///
    /// A real browser answers a fake `C:\fakepath\name` for security; there is
    /// no such boundary here — the program and the chooser are the same
    /// process — so the honest answer is the path the user actually picked.
    pub path: String,
    /// `multiple` — HTML §4.10.5.1.18.
    pub multiple: bool,
    pub disabled: bool,
    pub pressed: bool,
    pub colors: WidgetColors,
    pub width: f32,
    pub height: f32,
    pub id: WidgetId,
    pub name: String,
    pub font: crate::ide_text::FontSpec,
    rect: LayoutRect,
    pending_events: Vec<WidgetEvent>,
}

impl FileInput {
    pub fn new() -> Self {
        Self {
            path: String::new(),
            multiple: false,
            disabled: false,
            pressed: false,
            colors: WidgetColors {
                background: (239, 239, 239, 255),
                border: (118, 118, 118, 255),
                ..WidgetColors::default()
            },
            width: 240.0,
            height: 24.0,
            id: WidgetId::next(),
            name: String::new(),
            font: crate::ide_text::FontSpec::sans(13.0),
            rect: LayoutRect::zero(),
            pending_events: Vec::new(),
        }
    }

    pub fn with_name(mut self, name: &str) -> Self {
        self.name = name.to_string();
        self
    }

    /// The file name shown beside the button — the last path component, which
    /// is what a chooser displays and what a person recognises.
    fn display_name(&self) -> String {
        if self.path.is_empty() {
            return NOTHING.to_string();
        }
        self.path
            .rsplit(['/', '\\'])
            .next()
            .unwrap_or(&self.path)
            .to_string()
    }

    /// Open the platform chooser and keep what came back.
    ///
    /// Cancelling is not an error and not a change: `pick_file` answers `None`
    /// and the previous choice stands, which is what every file dialog does.
    fn choose(&mut self) {
        let dialog = crate::dialogs::FileDialog::new("Open");
        let picked = if self.multiple {
            dialog
                .open_multiple()
                .map(|paths| {
                    paths
                        .iter()
                        .map(|p| p.to_string_lossy().to_string())
                        .collect::<Vec<_>>()
                        .join(", ")
                })
                .filter(|joined| !joined.is_empty())
        } else {
            dialog.open().map(|p| p.to_string_lossy().to_string())
        };
        let Some(path) = picked else {
            return;
        };
        self.path = path;
        // `change` is the event HTML fires when a file is chosen. Reported as
        // `TextChanged` because that is this toolkit's word for "the value the
        // control holds is now this" — `dom.rs` maps it to `input`.
        self.pending_events
            .push(WidgetEvent::TextChanged(self.name.clone(), self.path.clone()));
    }
}

impl Default for FileInput {
    fn default() -> Self {
        Self::new()
    }
}

impl PanelWidget for FileInput {
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
        let ts = Transform::from_scale(ctx.scale, ctx.scale);
        let mut paint = Paint::default();
        paint.anti_alias = true;

        // The button half.
        let button_w = BUTTON_W.min(r.w);
        let (br, bg, bb, ba) = if self.pressed {
            (200, 200, 200, 255)
        } else {
            self.colors.background
        };
        paint.set_color_rgba8(br, bg, bb, ba);
        if let Some(path) = rounded_rect_path(r.x, r.y, button_w, r.h, 3.0) {
            ctx.pixmap.fill_path(&path, &paint, FillRule::Winding, ts, None);
        }
        let (er, eg, eb, ea) = self.colors.border;
        paint.set_color_rgba8(er, eg, eb, ea);
        if let Some(path) = rounded_rect_path(r.x, r.y, button_w, r.h, 3.0) {
            ctx.pixmap.stroke_path(
                &path,
                &paint,
                &Stroke {
                    width: 1.0,
                    ..Stroke::default()
                },
                ts,
                None,
            );
        }

        // Both labels sit on the control's centre line.
        let text_y = r.y + (r.h - self.font.size) / 2.0;
        let (fr, fg, fb, fa) = self.colors.foreground;
        ctx.draw_text(CHOOSE, r.x + 8.0, text_y, fr, fg, fb, fa);
        ctx.draw_text(
            &self.display_name(),
            r.x + button_w + 8.0,
            text_y,
            fr,
            fg,
            fb,
            fa,
        );
    }

    fn handle_mouse(&mut self, event: &MouseEvent) -> bool {
        if self.disabled || !self.rect.contains(event.x, event.y) {
            self.pressed = false;
            return false;
        }
        match event.kind {
            MouseEventKind::Press(LayoutMouseButton::Left) => {
                self.pressed = true;
                true
            }
            MouseEventKind::Release(LayoutMouseButton::Left) => {
                if self.pressed {
                    self.pressed = false;
                    // The WHOLE control opens the chooser, not just the button
                    // half — a browser does the same, and the filename is not
                    // separately clickable.
                    self.choose();
                }
                true
            }
            _ => false,
        }
    }

    fn handle_key(&mut self, _event: &KeyEvent) -> bool {
        false
    }

    fn handle_command(&mut self, cmd: &WidgetCommand) -> CommandValue {
        match cmd {
            // The value IS the path. A page may clear it (`value = ""`); HTML
            // does not let a page SET a real one, but refusing here would mean
            // a frontend that assigns a path silently gets nothing, and this
            // process has no privilege boundary to protect.
            WidgetCommand::SetText(t) => {
                self.path = t.clone();
                CommandValue::None
            }
            WidgetCommand::GetText | WidgetCommand::GetValue => {
                CommandValue::Text(self.path.clone())
            }
            WidgetCommand::SetEnabled(on) => {
                self.disabled = !on;
                CommandValue::None
            }
            WidgetCommand::Custom(name, value) if name == "SetMultiple" => {
                if let CommandValue::Text(v) = value {
                    self.multiple = !v.eq_ignore_ascii_case("false");
                }
                CommandValue::None
            }
            WidgetCommand::Custom(key, value) if self.font.apply_command(key, value) => {
                CommandValue::None
            }
            _ => CommandValue::None,
        }
    }

    fn drain_events(&mut self) -> Vec<WidgetEvent> {
        std::mem::take(&mut self.pending_events)
    }

    fn focusable(&self) -> bool {
        true
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    /// The value is the path, and an empty one reads back as empty rather than
    /// as the placeholder the control DRAWS.
    #[test]
    fn the_value_is_the_path_not_the_label() {
        let mut input = FileInput::new();
        // `CommandValue` has no `PartialEq` — matching is the read, and adding
        // a derive to a shared type for a test's convenience is not.
        let empty = input.handle_command(&WidgetCommand::GetValue);
        assert!(
            matches!(&empty, CommandValue::Text(v) if v.is_empty()),
            "nothing chosen is an empty value, not `No file chosen`: {empty:?}"
        );
        assert_eq!(input.display_name(), NOTHING);

        input.handle_command(&WidgetCommand::SetText("/tmp/report.pdf".into()));
        let chosen = input.handle_command(&WidgetCommand::GetValue);
        assert!(
            matches!(&chosen, CommandValue::Text(v) if v == "/tmp/report.pdf"),
            "the value is the whole path: {chosen:?}"
        );
        assert_eq!(
            input.display_name(),
            "report.pdf",
            "the control shows the file NAME; the value stays the whole path"
        );
    }

    /// Windows separators too — a VB program hands over `C:\Users\…` and the
    /// name is still the last component.
    #[test]
    fn a_windows_path_shows_its_last_component() {
        let mut input = FileInput::new();
        input.handle_command(&WidgetCommand::SetText("C:\\Users\\me\\notes.txt".into()));
        assert_eq!(input.display_name(), "notes.txt");
    }
}
