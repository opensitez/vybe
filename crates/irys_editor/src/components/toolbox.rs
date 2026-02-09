use dioxus::prelude::*;
use crate::app_state::AppState;
use irys_forms::ControlType;

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
    ];
    
    rsx! {
        div {
            class: "toolbox",
            style: "width: 150px; background: #fafafa; border-right: 1px solid #ccc; padding: 8px;",
            
            h3 { style: "margin: 0 0 8px 0; font-size: 14px;", "Toolbox" }
            
            div {
                style: "border-top: 1px solid #ccc; padding-top: 8px;",
                
                for (name, control_type) in controls {
                    {
                        let is_selected = *selected_tool.read() == control_type;
                        let bg_color = if is_selected { "#0078d4" } else { "transparent" };
                        let text_color = if is_selected { "white" } else { "black" };
                        
                        rsx! {
                            div {
                                key: "{name}",
                                style: "padding: 6px 8px; cursor: pointer; background: {bg_color}; color: {text_color}; border-radius: 3px; margin-bottom: 2px;",
                                onclick: move |_| {
                                    selected_tool.set(control_type);
                                },
                                "{name}"
                            }
                        }
                    }
                }
            }
        }
    }
}
