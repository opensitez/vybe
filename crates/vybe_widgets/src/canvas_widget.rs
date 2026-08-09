//! `Canvas` — a first-class drawable widget.
//!
//! Drop a `Canvas` into a Form like any other control. It owns a
//! [`RecordingCanvas`] that user code appends drawing commands to via
//! [`Canvas::with_canvas`] or [`Canvas::canvas_mut`]. The widget's
//! [`PanelWidget::render`] impl replays the recording onto the active
//! `tiny_skia::Pixmap` each frame.
//!
//! ## Why a separate widget vs. just `paint_overlay` on every widget
//!
//! Every `PanelWidget` gets a default-empty [`PanelWidget::paint_overlay`]
//! hook so user paint code can layer on top of any control. The `Canvas`
//! widget is the bare/empty version where there's no chrome to overlay
//! — it's pure drawing surface. Use it when you want a blank rectangle
//! the user owns entirely. Use `paint_overlay` on `Button` / `Label` /
//! etc. when you want to add custom decoration on top of standard
//! widget chrome.
//!
//! Both paths route through the same [`canvas::Canvas`] trait — the
//! difference is just where the recording lives.
//!
//! ## Standalone usage
//!
//! ```ignore
//! use vybe_widgets::{Form, Canvas as CanvasWidget};
//! use vybe_widgets::canvas::{Canvas, Color};
//!
//! let mut form = Form::new("Drawing");
//! let mut c = CanvasWidget::new()
//!     .with_name("art")
//!     .with_background(Color::rgb(255, 255, 255));
//!
//! c.with_canvas(|canvas| {
//!     canvas.set_fill_color(Color::rgb(220, 50, 50));
//!     canvas.fill_rect(10.0, 10.0, 100.0, 100.0);
//! });
//!
//! form.add_control(c, 0.0, 0.0, 800.0, 600.0);
//! ```
//!
//! Zero VM. Zero host. Zero .NET wrapper.

use crate::canvas::{Canvas as CanvasTrait, Color, RecordingCanvas, TinySkiaCanvas};
use crate::layout::{KeyEvent, LayoutRect, MouseEvent, PanelWidget, RenderContext, WidgetId};

/// A blank drawable surface widget.
///
/// Owns a `RecordingCanvas`. Replays it onto the active pixmap each
/// frame. Drawing persists across frames until the caller `clear()`s
/// the recording — matching .NET Graphics / VB6 / HTML5 canvas
/// semantics.
pub struct Canvas {
    name: String,
    rect: LayoutRect,
    id: WidgetId,
    background: Option<Color>,
    recording: RecordingCanvas,
    /// If true, the widget renders the recording in widget-relative
    /// coordinates by translating the canvas origin to the widget's
    /// top-left before replay. If false, the recording is in
    /// pixmap-absolute coordinates. Default: true.
    relative_coords: bool,
}

impl Canvas {
    /// Construct an empty canvas.
    pub fn new() -> Self {
        Self {
            name: String::new(),
            rect: LayoutRect::zero(),
            id: WidgetId::next(),
            background: None,
            recording: RecordingCanvas::new(),
            relative_coords: true,
        }
    }

    /// Set the widget's name (used for control lookup by the host
    /// bridge and the form's `send_command` API).
    pub fn with_name(mut self, name: &str) -> Self {
        self.name = name.to_string();
        self
    }

    /// Set the widget's background fill colour. `None` (the default)
    /// leaves the background transparent — whatever's underneath the
    /// widget shows through.
    pub fn with_background(mut self, color: Color) -> Self {
        self.background = Some(color);
        self
    }

    /// Switch between widget-relative and pixmap-absolute coordinate
    /// modes for the recording. Default: relative.
    pub fn with_relative_coords(mut self, relative: bool) -> Self {
        self.relative_coords = relative;
        self
    }

    /// Run a closure with mutable access to the underlying `Canvas`
    /// trait. Append commands; they're replayed on the next frame.
    pub fn with_canvas<F>(&mut self, f: F)
    where
        F: FnOnce(&mut dyn CanvasTrait),
    {
        f(&mut self.recording);
    }

    /// Direct mutable access to the underlying `RecordingCanvas`. Used
    /// by the host bridge (`vybe:gui::canvas*` host fns) which needs to
    /// hand the canvas trait to the host fn body.
    pub fn canvas_mut(&mut self) -> &mut RecordingCanvas {
        &mut self.recording
    }

    /// Snapshot of the captured commands. Used by tests.
    pub fn recording(&self) -> &RecordingCanvas {
        &self.recording
    }

    /// Clear all recorded commands. Drawing starts fresh on the next
    /// frame.
    pub fn clear(&mut self) {
        self.recording.clear();
    }
}

// Persistence semantics: drawings stay across frames until the caller
// invokes `clear()`. This matches retained-mode drawing models like
// HTML5 canvas — replay every frame, the recording is the source of
// truth.

impl Default for Canvas {
    fn default() -> Self {
        Self::new()
    }
}

impl PanelWidget for Canvas {
    fn name(&self) -> &str {
        &self.name
    }
    fn widget_id(&self) -> WidgetId {
        self.id
    }
    fn rect(&self) -> LayoutRect {
        self.rect
    }
    fn set_rect(&mut self, rect: LayoutRect) {
        self.rect = rect;
    }

    fn render(&mut self, ctx: &mut RenderContext) {
        if crate::canvas::trace_enabled() {
            let cmds = self.recording.commands_for_debug();
            let text_n = cmds
                .iter()
                .filter(|c| format!("{:?}", c).starts_with("FillText"))
                .count();
            eprintln!(
                "[canvas] RENDER widget={:?} total_cmds={} filltext={}",
                self.name,
                cmds.len(),
                text_n
            );
        }
        let r = self.rect;
        if r.w <= 0.0 || r.h <= 0.0 {
            return;
        }

        // 1. Background fill (if requested).
        if let Some(bg) = self.background {
            let mut paint = tiny_skia::Paint::default();
            paint.set_color_rgba8(bg.r, bg.g, bg.b, bg.a);
            paint.anti_alias = true;
            if let Some(rect) = tiny_skia::Rect::from_xywh(
                r.x * ctx.scale,
                r.y * ctx.scale,
                r.w * ctx.scale,
                r.h * ctx.scale,
            ) {
                ctx.pixmap
                    .fill_rect(rect, &paint, tiny_skia::Transform::identity(), None);
            }
        }

        // 2. Replay the recording onto a text-enabled TinySkiaCanvas.
        //
        // The canvas is constructed with `with_text(...)` so text
        // commands captured in the recording (FillText/StrokeText)
        // render through cosmic-text. The widget's origin is
        // pre-applied as a translate so user-supplied coordinates are
        // widget-relative. (If the user opted out via
        // `with_relative_coords(false)`, the recording is in absolute
        // pixmap coordinates and we leave the transform at identity.)
        if !self.recording.is_empty() {
            let mut canvas =
                TinySkiaCanvas::with_text(&mut ctx.pixmap, ctx.font_system, ctx.swash_cache);
            if self.relative_coords {
                canvas.translate(r.x * ctx.scale, r.y * ctx.scale);
            }
            if (ctx.scale - 1.0).abs() > f32::EPSILON {
                canvas.scale(ctx.scale, ctx.scale);
            }
            self.recording.replay(&mut canvas);
        }
    }

    fn handle_mouse(&mut self, _event: &MouseEvent) -> bool {
        false
    }
    fn handle_key(&mut self, _event: &KeyEvent) -> bool {
        false
    }

    /// Used by the host bridge to find this widget by name and pull
    /// out its underlying `RecordingCanvas`. See
    /// `vybe_host::gui_state::GuiState::find_canvas_mut`.
    fn as_any_mut(&mut self) -> Option<&mut dyn std::any::Any> {
        Some(self as &mut dyn std::any::Any)
    }
}
