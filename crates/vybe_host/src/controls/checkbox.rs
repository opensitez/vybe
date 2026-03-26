use super::{ControlDef, Props, base_css};

pub static DEF: ControlDef = ControlDef {
    tag: "label",
    inner_tag: Some("input"),
    props: &["Text", "Checked", "Enabled", "Visible", "BackColor", "ForeColor"],
    events: &["CheckedChanged", "Click"],
    default_size: (100, 24),
    css_fn: css,
    container: false,
    input_type: Some("checkbox"),
    extra_attrs: &[],
};

fn css(props: &Props) -> String {
    let mut s = String::from("display: flex; align-items: center; gap: 6px; cursor: pointer; user-select: none; ");
    s.push_str(&base_css(props));
    s
}
