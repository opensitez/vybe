use super::{ControlDef, Props, base_css};

pub static DEF: ControlDef = ControlDef {
    tag: "div",
    inner_tag: None,
    props: &["Text", "Visible", "BackColor", "ForeColor", "Font", "TextAlign", "AutoSize"],
    events: &["Click"],
    default_size: (100, 23),
    css_fn: css,
    container: false,
    input_type: None,
    extra_attrs: &[],
};

pub static LINK_DEF: ControlDef = ControlDef {
    tag: "a",
    inner_tag: None,
    props: &["Text", "Visible", "ForeColor", "Font", "LinkColor"],
    events: &["Click", "LinkClicked"],
    default_size: (100, 23),
    css_fn: link_css,
    container: false,
    input_type: None,
    extra_attrs: &[("href", "#")],
};

fn css(props: &Props) -> String {
    let mut s = String::from("display: flex; align-items: center; user-select: none; ");
    if let Some(align) = props.get("TextAlign") {
        match align.as_str() {
            "MiddleCenter" | "TopCenter" | "BottomCenter" => s.push_str("justify-content: center; "),
            "MiddleRight" | "TopRight" | "BottomRight" => s.push_str("justify-content: flex-end; "),
            _ => {}
        }
    }
    s.push_str(&base_css(props));
    s
}

fn link_css(props: &Props) -> String {
    let mut s = String::from("display: flex; align-items: center; color: #0066cc; text-decoration: underline; cursor: pointer; ");
    s.push_str(&base_css(props));
    s
}
