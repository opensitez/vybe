use super::{ControlDef, Props, base_css};

pub static DEF: ControlDef = ControlDef {
    tag: "button",
    inner_tag: None,
    props: &["Text", "Enabled", "Visible", "BackColor", "ForeColor", "Font", "FlatStyle"],
    events: &["Click", "MouseEnter", "MouseLeave", "MouseDown", "MouseUp"],
    default_size: (100, 30),
    css_fn: css,
    container: false,
    input_type: None,
    extra_attrs: &[],
};

fn css(props: &Props) -> String {
    let flat = props.get("FlatStyle").map(|v| v == "Flat").unwrap_or(false);
    let mut s = String::from("cursor: pointer; padding: 4px 12px; border: 1px solid #999; border-radius: 3px; ");
    if flat {
        s.push_str("background: transparent; border: 1px solid #ccc; ");
    } else {
        s.push_str("background: #e0e0e0; ");
    }
    s.push_str(&base_css(props));
    s
}
