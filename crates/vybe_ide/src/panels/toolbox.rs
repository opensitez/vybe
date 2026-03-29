use egui::Ui;
use vybe_forms::ControlType;
use crate::state::EditorState;

const VISUAL_CONTROLS: &[(&str, Option<ControlType>)] = &[
    ("Pointer",           None),
    ("Button",            Some(ControlType::Button)),
    ("Label",             Some(ControlType::Label)),
    ("LinkLabel",         Some(ControlType::LinkLabel)),
    ("TextBox",           Some(ControlType::TextBox)),
    ("MaskedTextBox",     Some(ControlType::MaskedTextBox)),
    ("RichTextBox",       Some(ControlType::RichTextBox)),
    ("CheckBox",          Some(ControlType::CheckBox)),
    ("RadioButton",       Some(ControlType::RadioButton)),
    ("ComboBox",          Some(ControlType::ComboBox)),
    ("ListBox",           Some(ControlType::ListBox)),
    ("Panel",             Some(ControlType::Panel)),
    ("GroupBox",          Some(ControlType::Frame)),
    ("PictureBox",        Some(ControlType::PictureBox)),
    ("DataGridView",      Some(ControlType::DataGridView)),
    ("TreeView",          Some(ControlType::TreeView)),
    ("ListView",          Some(ControlType::ListView)),
    ("TabControl",        Some(ControlType::TabControl)),
    ("ProgressBar",       Some(ControlType::ProgressBar)),
    ("TrackBar",          Some(ControlType::TrackBar)),
    ("NumericUpDown",     Some(ControlType::NumericUpDown)),
    ("DateTimePicker",    Some(ControlType::DateTimePicker)),
    ("MonthCalendar",     Some(ControlType::MonthCalendar)),
    ("WebBrowser",        Some(ControlType::WebBrowser)),
    ("MenuStrip",         Some(ControlType::MenuStrip)),
    ("ContextMenuStrip",  Some(ControlType::ContextMenuStrip)),
    ("StatusStrip",       Some(ControlType::StatusStrip)),
    ("ToolStrip",         Some(ControlType::ToolStrip)),
    ("SplitContainer",    Some(ControlType::SplitContainer)),
    ("FlowLayoutPanel",   Some(ControlType::FlowLayoutPanel)),
    ("TableLayoutPanel",  Some(ControlType::TableLayoutPanel)),
    ("HScrollBar",        Some(ControlType::HScrollBar)),
    ("VScrollBar",        Some(ControlType::VScrollBar)),
];

/// (label, icon, control_type)
/// Non-visual controls are added immediately on click — no canvas interaction needed.
const DATA_CONTROLS: &[(&str, &str, ControlType)] = &[
    ("BindingSource",      "🔗", ControlType::BindingSourceComponent),
    ("BindingNavigator",   "🧭", ControlType::BindingNavigator),
    ("DataSet",            "🗄",  ControlType::DataSetComponent),
    ("DataTable",          "📋", ControlType::DataTableComponent),
    ("DataAdapter",        "🔌", ControlType::DataAdapterComponent),
    ("Timer",              "⏱",  ControlType::Timer),
    ("ImageList",          "🖼",  ControlType::ImageList),
    ("ErrorProvider",      "⚠",  ControlType::ErrorProvider),
    ("BackgroundWorker",   "⚙",  ControlType::BackgroundWorker),
    ("OpenFileDialog",     "📂", ControlType::OpenFileDialog),
    ("SaveFileDialog",     "💾", ControlType::SaveFileDialog),
    ("FolderBrowserDialog","📁", ControlType::FolderBrowserDialog),
    ("FontDialog",         "🔤", ControlType::FontDialog),
    ("ColorDialog",        "🎨", ControlType::ColorDialog),
    ("NotifyIcon",         "🔔", ControlType::NotifyIcon),
];

pub fn show(ui: &mut Ui, state: &mut EditorState) {
    ui.heading("Toolbox");
    ui.separator();

    // ── Standard Controls ─────────────────────────────────────────────────────
    for (label, ct) in VISUAL_CONTROLS {
        let selected = &state.selected_tool == ct;
        if ui.selectable_label(selected, *label).clicked() {
            state.selected_tool = ct.clone();
        }
    }

    ui.add_space(8.0);
    ui.label(egui::RichText::new("Data").small().weak());
    ui.separator();

    // ── Data / Non-visual Controls ────────────────────────────────────────────
    // These are added immediately on click — no canvas placement step.
    for (label, icon, ct) in DATA_CONTROLS {
        let is_non_visual = ct.is_non_visual();
        let display = format!("{} {}", icon, label);
        let selected = state.selected_tool.as_ref() == Some(ct);
        if ui.selectable_label(selected, display).clicked() {
            if is_non_visual {
                // Add directly at (0,0) — will appear in the component tray
                state.add_control(ct.clone(), 0, 0);
                state.selected_tool = None;
            } else {
                state.selected_tool = Some(ct.clone());
            }
        }
    }
}
