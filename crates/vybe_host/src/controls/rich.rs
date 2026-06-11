use super::{ControlDef, Props, base_css};

pub static RICHTEXTBOX_DEF: ControlDef = ControlDef {
    tag: "div",
    inner_tag: None,
    props: &[
        "Text",
        "Rtf",
        "ReadOnly",
        "Enabled",
        "Visible",
        "BackColor",
        "ForeColor",
    ],
    events: &["TextChanged", "LinkClicked"],
    default_size: (100, 96),
    css_fn: rich_css,
    container: false,
    input_type: None,
    extra_attrs: &[("contenteditable", "true")],
};

pub static WEBBROWSER_DEF: ControlDef = ControlDef {
    tag: "iframe",
    inner_tag: None,
    props: &["Url", "DocumentText", "Visible"],
    events: &["DocumentCompleted", "Navigating"],
    default_size: (240, 150),
    css_fn: web_css,
    container: false,
    input_type: None,
    extra_attrs: &[("sandbox", "allow-scripts")],
};

pub static PICTUREBOX_DEF: ControlDef = ControlDef {
    tag: "img",
    inner_tag: None,
    props: &[
        "Image",
        "ImageLocation",
        "SizeMode",
        "Visible",
        "BackColor",
        "BorderStyle",
    ],
    events: &["Click", "Paint"],
    default_size: (100, 50),
    css_fn: picture_css,
    container: false,
    input_type: None,
    extra_attrs: &[],
};

fn rich_css(props: &Props) -> String {
    let mut s = String::from(
        "border: 1px solid #999; padding: 4px; overflow: auto; white-space: pre-wrap; ",
    );
    if props.get("ReadOnly").map(|v| v == "True").unwrap_or(false) {
        s.push_str("background: #f8f8f8; ");
    }
    s.push_str(&base_css(props));
    s
}

fn web_css(props: &Props) -> String {
    let mut s = String::from("border: 1px solid #ccc; ");
    s.push_str(&base_css(props));
    s
}

fn picture_css(props: &Props) -> String {
    let mut s = String::from("object-fit: contain; ");
    match props.get("SizeMode").map(|v| v.as_str()) {
        Some("StretchImage") => s.push_str("object-fit: fill; "),
        Some("Zoom") => s.push_str("object-fit: contain; "),
        Some("CenterImage") => s.push_str("object-fit: none; object-position: center; "),
        _ => {}
    }
    s.push_str(&base_css(props));
    s
}
