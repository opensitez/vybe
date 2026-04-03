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

pub mod checkbox;
pub mod radio;
pub mod textfield;
pub mod select;
pub mod dropdown;
pub mod button;
pub mod slider;
pub mod tree_view;
pub mod tabs;
pub mod progress;
pub mod grid;
pub mod label;
pub mod listbox;
pub mod panel;
pub mod picturebox;
pub mod list_view;
pub mod numeric;
pub mod datetime;
pub mod scrollbar;
pub mod link_label;
pub mod masked_textbox;
pub mod groupbox;
pub mod month_calendar;
pub mod menu_strip;
pub mod context_menu;
pub mod status_strip;
pub mod tool_strip;
pub mod split_container;
pub mod flow_layout;
pub mod table_layout;
pub mod toolbox;
pub mod color_picker;
pub mod font_picker;
pub mod language;
pub mod ide_text;
pub mod text_editor;
pub mod code_editor_widget;
pub mod resource_editor;
pub mod layout;
pub mod dock_panel;
pub mod split_panel;
pub mod tab_panel;
pub mod status_bar_panel;
pub mod app_window;

pub use checkbox::Checkbox;
pub use radio::Radio;
pub use textfield::TextInput;
pub use select::Select;
pub use dropdown::{Dropdown, DropdownEvent};
pub use button::Button;
pub use slider::Slider;
pub use tree_view::{TreeView, TreeEvent, FileEntry};
pub use tabs::Tabs;
pub use progress::ProgressBar;
pub use grid::DataGrid;
pub use label::Label;
pub use listbox::ListBox;
pub use panel::{Panel, BorderStyle};
pub use picturebox::PictureBox;
pub use list_view::ListView;
pub use numeric::NumericUpDown;
pub use datetime::DateTimePicker;
pub use scrollbar::ScrollBar;
pub use link_label::LinkLabel;
pub use masked_textbox::MaskedTextBox;
pub use groupbox::GroupBox;
pub use month_calendar::MonthCalendar;
pub use menu_strip::MenuStrip;
pub use context_menu::ContextMenu;
pub use status_strip::StatusStrip;
pub use tool_strip::{ToolStrip, ToolStripItem};
pub use split_container::SplitContainer;
pub use flow_layout::FlowLayoutPanel;
pub use table_layout::TableLayoutPanel;
pub use toolbox::Toolbox;
pub use color_picker::{ColorPicker, PickedColor, ColorPickerEvent};
pub use font_picker::{FontPicker, FontPickerEvent};
pub use language::{LanguageDef, CommentDef, load_language};
pub use ide_text::{draw_text, measure_text};
pub use text_editor::{TextEditor, TokenKind, TokenSpan, LexerState, DiagnosticInfo, DiagnosticSeverity};
pub use code_editor_widget::{CodeEditorWidget, Theme};
pub use resource_editor::{ResourceEditor, ResourceEditorEvent, ResourceEntry, ResourceTab};

// ── GUI Toolkit ────────────────────────────────────────────────────────
pub use layout::{LayoutRect, MouseButton, MouseEventKind, MouseEvent, KeyEvent, RenderContext, Dock, PanelWidget};
pub use dock_panel::DockPanel;
pub use split_panel::SplitPanel;
pub use tab_panel::TabPanel;
pub use status_bar_panel::StatusBarPanel;
pub use app_window::{Application, run_app};

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
