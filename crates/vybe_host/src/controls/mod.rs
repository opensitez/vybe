//! Shared control definitions — data model for all UI controls.
//! Each control defines its tag, properties, events, default size, and CSS.
//! The renderer reads these definitions — no rendering logic here.

pub mod button;
pub mod checkbox;
pub mod combobox;
pub mod datetime;
pub mod dialog;
pub mod grid;
pub mod label;
pub mod listbox;
pub mod menu;
pub mod numeric;
pub mod panel;
pub mod progress;
pub mod radio;
pub mod rich;
pub mod tabs;
pub mod textbox;

use std::collections::HashMap;

/// Properties passed to a control's CSS/render function.
pub type Props = HashMap<String, String>;

/// A control definition — describes how to render a control type.
pub struct ControlDef {
    /// HTML tag to use ("button", "input", "div", "select", etc.)
    pub tag: &'static str,
    /// Inner HTML structure (for complex controls). Empty = just use text content.
    pub inner_tag: Option<&'static str>,
    /// Supported property names
    pub props: &'static [&'static str],
    /// Events this control can fire
    pub events: &'static [&'static str],
    /// Default width, height
    pub default_size: (i32, i32),
    /// CSS generation function: takes properties, returns CSS string
    pub css_fn: fn(&Props) -> String,
    /// Whether this control is a container (can have children)
    pub container: bool,
    /// HTML input type (for input-based controls)
    pub input_type: Option<&'static str>,
    /// Extra HTML attributes
    pub extra_attrs: &'static [(&'static str, &'static str)],
}

/// Get the control definition for a control type name (case-insensitive).
pub fn get_def(name: &str) -> &'static ControlDef {
    match name.to_lowercase().as_str() {
        "button" => &button::DEF,
        "label" => &label::DEF,
        "textbox" => &textbox::DEF,
        "checkbox" => &checkbox::DEF,
        "radiobutton" => &radio::DEF,
        "combobox" => &combobox::DEF,
        "listbox" => &listbox::DEF,
        "panel" | "groupbox" | "frame" => &panel::DEF,
        "progressbar" => &progress::PROGRESS_DEF,
        "trackbar" => &progress::TRACKBAR_DEF,
        "datagridview" => &grid::DATAGRID_DEF,
        "listview" => &grid::LISTVIEW_DEF,
        "tabcontrol" => &tabs::TABCONTROL_DEF,
        "tabpage" => &tabs::TABPAGE_DEF,
        "menustrip" => &menu::MENUSTRIP_DEF,
        "contextmenustrip" => &menu::CONTEXTMENU_DEF,
        "toolstrip" => &menu::TOOLSTRIP_DEF,
        "datetimepicker" => &datetime::DATETIME_DEF,
        "monthcalendar" => &datetime::CALENDAR_DEF,
        "numericupdown" => &numeric::DEF,
        "richtextbox" => &rich::RICHTEXTBOX_DEF,
        "webbrowser" => &rich::WEBBROWSER_DEF,
        "picturebox" => &rich::PICTUREBOX_DEF,
        "linklabel" => &label::LINK_DEF,
        "statusstrip" => &menu::STATUSSTRIP_DEF,
        "maskedtextbox" => &textbox::MASKED_DEF,
        "splitcontainer" => &panel::SPLIT_DEF,
        "flowlayoutpanel" => &panel::FLOW_DEF,
        "tablelayoutpanel" => &panel::TABLE_DEF,
        "hscrollbar" | "vscrollbar" => &progress::SCROLLBAR_DEF,
        "bindingnavigator" => &menu::TOOLSTRIP_DEF,
        _ => &label::DEF, // fallback: render as label
    }
}

/// Base CSS that all controls share.
pub fn base_css(props: &Props) -> String {
    let mut css = String::new();
    if let Some(bg) = props.get("BackColor") {
        css.push_str(&format!("background-color: {};", bg));
    }
    if let Some(fg) = props.get("ForeColor") {
        css.push_str(&format!("color: {};", fg));
    }
    if let Some(font) = props.get("Font") {
        css.push_str(&format!("font: {};", font));
    }
    if props.get("Enabled").map(|v| v == "false").unwrap_or(false) {
        css.push_str("opacity: 0.5; pointer-events: none;");
    }
    if props.get("Visible").map(|v| v == "false").unwrap_or(false) {
        css.push_str("display: none;");
    }
    css
}
