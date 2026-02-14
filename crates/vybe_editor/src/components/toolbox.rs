use dioxus::prelude::*;
use crate::app_state::AppState;
use vybe_forms::ControlType;

#[component]
pub fn Toolbox() -> Element {
    let state = use_context::<AppState>();
    let mut selected_tool = state.selected_tool;
    
    let controls = vec![
        ("Pointer", None),
        ("Button", Some(ControlType::Button)),
        ("Label", Some(ControlType::Label)),
        ("TextBox", Some(ControlType::TextBox)),
        ("CheckBox", Some(ControlType::CheckBox)),
        ("RadioButton", Some(ControlType::RadioButton)),
        ("ComboBox", Some(ControlType::ComboBox)),
        ("ListBox", Some(ControlType::ListBox)),
        ("Frame", Some(ControlType::Frame)),
        ("PictureBox", Some(ControlType::PictureBox)),
        ("RichTextBox", Some(ControlType::RichTextBox)),
        ("WebBrowser", Some(ControlType::WebBrowser)),
        ("TreeView", Some(ControlType::TreeView)),
        ("DataGridView", Some(ControlType::DataGridView)),
        ("Panel", Some(ControlType::Panel)),
        ("ListView", Some(ControlType::ListView)),
        ("TabControl", Some(ControlType::TabControl)),
        ("ProgressBar", Some(ControlType::ProgressBar)),
        ("NumericUpDown", Some(ControlType::NumericUpDown)),
        ("MenuStrip", Some(ControlType::MenuStrip)),
        ("ContextMenuStrip", Some(ControlType::ContextMenuStrip)),
        ("StatusStrip", Some(ControlType::StatusStrip)),
        ("DateTimePicker", Some(ControlType::DateTimePicker)),
        ("LinkLabel", Some(ControlType::LinkLabel)),
        ("ToolStrip", Some(ControlType::ToolStrip)),
        ("TrackBar", Some(ControlType::TrackBar)),
        ("MaskedTextBox", Some(ControlType::MaskedTextBox)),
        ("SplitContainer", Some(ControlType::SplitContainer)),
        ("FlowLayoutPanel", Some(ControlType::FlowLayoutPanel)),
        ("TableLayoutPanel", Some(ControlType::TableLayoutPanel)),
        ("MonthCalendar", Some(ControlType::MonthCalendar)),
        ("HScrollBar", Some(ControlType::HScrollBar)),
        ("VScrollBar", Some(ControlType::VScrollBar)),
    ];

    let data_controls = vec![
        ("BindingSource", Some(ControlType::BindingSourceComponent)),
        ("BindingNavigator", Some(ControlType::BindingNavigator)),
        ("DataSet", Some(ControlType::DataSetComponent)),
        ("DataTable", Some(ControlType::DataTableComponent)),
        ("DataAdapter", Some(ControlType::DataAdapterComponent)),
    ];
    
    rsx! {
        div {
            class: "toolbox",
            style: "width: 150px; background: #fafafa; border-right: 1px solid #ccc; padding: 8px; overflow-y: auto;",
            
            h3 { style: "margin: 0 0 8px 0; font-size: 14px;", "Toolbox" }
            
            // Standard controls
            div {
                style: "border-top: 1px solid #ccc; padding-top: 8px;",
                
                for (name, control_type) in controls {
                    {
                        let is_selected = *selected_tool.read() == control_type;
                        let bg_color = if is_selected { "#0078d4" } else { "transparent" };
                        let text_color = if is_selected { "white" } else { "black" };
                        let ct_click = control_type.clone();
                        
                        rsx! {
                            div {
                                key: "{name}",
                                style: "padding: 6px 8px; cursor: pointer; background: {bg_color}; color: {text_color}; border-radius: 3px; margin-bottom: 2px;",
                                onclick: move |_| {
                                    selected_tool.set(ct_click.clone());
                                },
                                "{name}"
                            }
                        }
                    }
                }
            }

            // Data section
            h4 { style: "margin: 12px 0 4px 0; font-size: 12px; color: #666; text-transform: uppercase; letter-spacing: 0.5px;", "Data" }
            div {
                style: "border-top: 1px solid #ccc; padding-top: 4px;",

                for (name, control_type) in data_controls {
                    {
                        let is_selected = *selected_tool.read() == control_type;
                        let bg_color = if is_selected { "#0078d4" } else { "transparent" };
                        let text_color = if is_selected { "white" } else { "black" };
                        let icon = match name {
                            "BindingSource" => "🔗 ",
                            "BindingNavigator" => "🧭 ",
                            "DataSet" => "🗄️ ",
                            "DataTable" => "📋 ",
                            "DataAdapter" => "🔌 ",
                            _ => "",
                        };
                        let ct_click = control_type.clone();

                        rsx! {
                            div {
                                key: "{name}",
                                style: "padding: 5px 8px; cursor: pointer; background: {bg_color}; color: {text_color}; border-radius: 3px; margin-bottom: 2px; font-size: 12px;",
                                onclick: move |_| {
                                    // Non-visual components are added immediately (no form canvas click needed)
                                    if let Some(ct) = &ct_click {
                                        if ct.is_non_visual() {
                                            state.add_control_at(ct.clone(), 0, 0);
                                            selected_tool.set(None); // reset to pointer
                                            return;
                                        }
                                    }
                                    selected_tool.set(ct_click.clone());
                                },
                                "{icon}{name}"
                            }
                        }
                    }
                }
            }
        }
    }
}
