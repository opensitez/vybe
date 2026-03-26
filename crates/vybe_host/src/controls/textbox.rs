use super::{ControlDef, Props, base_css};

pub static DEF: ControlDef = ControlDef {
    tag: "input",
    inner_tag: None,
    props: &["Text", "Enabled", "Visible", "BackColor", "ForeColor", "Font", "ReadOnly", "MaxLength", "PasswordChar", "Multiline"],
    events: &["TextChanged", "KeyPress", "KeyDown", "KeyUp", "GotFocus", "LostFocus"],
    default_size: (100, 23),
    css_fn: css,
    container: false,
    input_type: Some("text"),
    extra_attrs: &[],
};

pub static MASKED_DEF: ControlDef = ControlDef {
    tag: "input",
    inner_tag: None,
    props: &["Text", "Mask", "Enabled", "Visible"],
    events: &["TextChanged", "GotFocus", "LostFocus"],
    default_size: (100, 23),
    css_fn: css,
    container: false,
    input_type: Some("text"),
    extra_attrs: &[],
};

fn css(props: &Props) -> String {
    let mut s = String::from("padding: 2px 4px; border: 1px solid #999; box-sizing: border-box; ");
    if props.get("Multiline").map(|v| v == "True").unwrap_or(false) {
        s.push_str("resize: vertical; ");
    }
    if props.get("ReadOnly").map(|v| v == "True").unwrap_or(false) {
        s.push_str("background: #f0f0f0; ");
    }
    s.push_str(&base_css(props));
    s
}
