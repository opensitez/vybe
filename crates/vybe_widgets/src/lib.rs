//! Vibe Widgets — standalone tiny-skia GUI widgets.
//!
//! These widgets render form elements using only tiny-skia. They have zero
//! dependency on the HTML/CSS/DOM engine and can be used standalone:
//!
//! ```ignore
//! use vibe_widgets::{Checkbox, TextInput, Widget};
//!
//! let mut checkbox = Checkbox::new("Accept terms");
//! checkbox.paint(&mut pixmap, 10.0, 10.0, 1.0);
//! if checkbox.click(mouse_x, mouse_y) { /* toggled */ }
//! ```
//!
//! Inside the browser engine, these same widgets are used to render HTML form
//! elements (`<input>`, `<select>`, `<textarea>`, etc.).

pub mod app_window;
pub mod binding_navigator;
pub mod button;
pub mod canvas;
pub mod canvas_widget;
pub mod checkbox;
pub mod code_editor_widget;
pub mod color_picker;
pub mod context_menu;
pub mod datetime;
pub mod dialogs;
pub mod dock_panel;
pub mod dropdown;
pub mod flow_layout;
pub mod font_picker;
pub mod form;
pub mod grid;
pub mod groupbox;
pub mod ide_text;
pub mod label;
pub mod language;
pub mod layout;
pub mod link_label;
pub mod list_view;
pub mod listbox;
pub mod masked_textbox;
pub mod menu_strip;
pub mod month_calendar;
pub mod numeric;
pub mod output_panel;
pub mod panel;
pub mod picturebox;
pub mod progress;
pub mod properties_panel;
pub mod radio;
pub mod resource_editor;
pub mod scrollbar;
pub mod select;
pub mod slider;
pub mod split_container;
pub mod split_panel;
pub mod stack_panel;
pub mod status_bar_panel;
pub mod status_strip;
pub mod tab_panel;
pub mod table_layout;
pub mod tabs;
pub mod text_editor;
pub mod textfield;
pub mod tool_strip;
pub mod toolbox;
pub mod tree_view;
pub mod wrap_panel;

pub use button::Button;
pub use checkbox::Checkbox;
pub use code_editor_widget::{CodeEditorWidget, Theme};
pub use color_picker::{ColorPicker, ColorPickerEvent, PickedColor};
pub use context_menu::ContextMenu;
pub use datetime::DateTimePicker;
pub use dropdown::{Dropdown, DropdownEvent};
pub use flow_layout::{FlowDirection, FlowLayoutPanel};
pub use font_picker::{FontPicker, FontPickerEvent};
pub use grid::DataGrid;
pub use groupbox::GroupBox;
pub use ide_text::{draw_text, measure_text};
pub use label::Label;
pub use language::{CommentDef, LanguageDef, load_language};
pub use link_label::LinkLabel;
pub use list_view::ListView;
pub use listbox::ListBox;
pub use masked_textbox::MaskedTextBox;
pub use menu_strip::MenuStrip;
pub use month_calendar::MonthCalendar;
pub use numeric::NumericUpDown;
pub use panel::{BorderStyle, Panel};
pub use picturebox::PictureBox;
pub use progress::ProgressBar;
pub use radio::Radio;
pub use resource_editor::{ResourceEditor, ResourceEditorEvent, ResourceEntry, ResourceTab};
pub use scrollbar::ScrollBar;
pub use select::Select;
pub use slider::Slider;
pub use split_container::SplitContainer;
pub use status_strip::StatusStrip;
pub use table_layout::{SizeMode, TableLayoutPanel};
pub use tabs::Tabs;
pub use text_editor::{
    DiagnosticInfo, DiagnosticSeverity, LexerState, TextEditor, TokenKind, TokenSpan,
};
pub use textfield::TextInput;
pub use tool_strip::{ToolStrip, ToolStripItem};
pub use toolbox::Toolbox;
pub use tree_view::{FileEntry, TreeEvent, TreeView};

// ── GUI Toolkit ────────────────────────────────────────────────────────
pub use app_window::{Application, run_app};
pub use binding_navigator::{BindingNavigator, NavAction};
pub use canvas_widget::Canvas;
pub use dock_panel::DockPanel;
pub use form::Form;
pub use layout::{
    Anchor, AnchorLayout, CheckState, CommandValue, CursorMotion, Dock, FocusManager, KeyEvent,
    LayoutRect, MouseButton, MouseEvent, MouseEventKind, NullWidget, PanelWidget, RenderContext,
    SelectionMode, TextAlign, WidgetCommand, WidgetEvent, WidgetId, apply_anchor_layouts,
};
pub use output_panel::{OutputPanel, OutputPanelEvent, OutputTab, ProblemEntry, ProblemSeverity};
pub use properties_panel::{PropEvent, PropItem, PropTab, PropertiesPanel};
pub use split_panel::SplitPanel;
pub use stack_panel::{Orientation, StackPanel};
pub use status_bar_panel::StatusBarPanel;
pub use tab_panel::TabPanel;
pub use winit::window::CursorIcon;
/// Re-exported so consumers (vybex's FormApp) can NAME the key/state types
/// that already appear in `KeyEvent`'s public fields without depending on
/// winit directly.
pub use winit;
pub use wrap_panel::WrapPanel;

// ── Re-exports from cosmic-text (so consumers don't need cosmic_text directly) ──
pub use cosmic_text::{Color as TextColor, FontSystem, SwashCache};

// ── Re-exports from tiny-skia (so consumers don't need tiny-skia directly) ──
pub use tiny_skia::Pixmap;

/// Fill a pixmap with a solid RGBA background colour.
pub fn fill_background(pixmap: &mut tiny_skia::Pixmap, r: u8, g: u8, b: u8, a: u8) {
    pixmap.fill(tiny_skia::Color::from_rgba8(r, g, b, a));
}

use tiny_skia::PathBuilder;

/// Colors used by widgets.
#[derive(Clone, Copy, Debug)]
pub struct WidgetColors {
    pub foreground: (u8, u8, u8, u8),
    pub background: (u8, u8, u8, u8),
    pub border: (u8, u8, u8, u8),
    pub accent: (u8, u8, u8, u8),
    pub placeholder: (u8, u8, u8, u8),
    pub focus_ring: (u8, u8, u8, u8),
}

impl Default for WidgetColors {
    fn default() -> Self {
        Self {
            foreground: (51, 51, 51, 255),
            background: (255, 255, 255, 255),
            border: (118, 118, 118, 255),
            accent: (0, 102, 204, 255),
            placeholder: (128, 128, 128, 128),
            focus_ring: (0, 102, 204, 128),
        }
    }
}

/// Shared helper: draw a rounded rectangle path.
pub fn rounded_rect_path(x: f32, y: f32, w: f32, h: f32, r: f32) -> Option<tiny_skia::Path> {
    let r = r.min(w / 2.0).min(h / 2.0);
    let mut pb = PathBuilder::new();
    pb.move_to(x + r, y);
    pb.line_to(x + w - r, y);
    pb.quad_to(x + w, y, x + w, y + r);
    pb.line_to(x + w, y + h - r);
    pb.quad_to(x + w, y + h, x + w - r, y + h);
    pb.line_to(x + r, y + h);
    pb.quad_to(x, y + h, x, y + h - r);
    pb.line_to(x, y + r);
    pb.quad_to(x, y, x + r, y);
    pb.close();
    pb.finish()
}

/// Shared helper: draw a circle path using 4 quadratic bezier arcs.
pub fn circle_path(cx: f32, cy: f32, r: f32) -> Option<tiny_skia::Path> {
    let k = r * 0.5522848; // magic number for circular arc approximation
    let mut pb = PathBuilder::new();
    pb.move_to(cx, cy - r);
    pb.cubic_to(cx + k, cy - r, cx + r, cy - k, cx + r, cy);
    pb.cubic_to(cx + r, cy + k, cx + k, cy + r, cx, cy + r);
    pb.cubic_to(cx - k, cy + r, cx - r, cy + k, cx - r, cy);
    pb.cubic_to(cx - r, cy - k, cx - k, cy - r, cx, cy - r);
    pb.close();
    pb.finish()
}
