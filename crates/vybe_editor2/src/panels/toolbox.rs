use egui::Ui;
use vybe_forms::ControlType;
use crate::state::EditorState;

const CONTROLS: &[(&str, Option<ControlType>)] = &[
    ("Pointer",         None),
    ("Button",          Some(ControlType::Button)),
    ("Label",           Some(ControlType::Label)),
    ("TextBox",         Some(ControlType::TextBox)),
    ("CheckBox",        Some(ControlType::CheckBox)),
    ("RadioButton",     Some(ControlType::RadioButton)),
    ("ComboBox",        Some(ControlType::ComboBox)),
    ("ListBox",         Some(ControlType::ListBox)),
    ("Panel",           Some(ControlType::Panel)),
    ("GroupBox",        Some(ControlType::Frame)),
    ("PictureBox",      Some(ControlType::PictureBox)),
    ("DataGridView",    Some(ControlType::DataGridView)),
    ("TreeView",        Some(ControlType::TreeView)),
    ("ListView",        Some(ControlType::ListView)),
    ("ProgressBar",     Some(ControlType::ProgressBar)),
    ("TabControl",      Some(ControlType::TabControl)),
    ("DateTimePicker",  Some(ControlType::DateTimePicker)),
    ("NumericUpDown",   Some(ControlType::NumericUpDown)),
    ("TrackBar",        Some(ControlType::TrackBar)),
    ("RichTextBox",     Some(ControlType::RichTextBox)),
];

pub fn show(ui: &mut Ui, state: &mut EditorState) {
    ui.heading("Toolbox");
    ui.separator();
    for (label, ct) in CONTROLS {
        let selected = &state.selected_tool == ct;
        if ui.selectable_label(selected, *label).clicked() {
            state.selected_tool = ct.clone();
        }
    }
}
