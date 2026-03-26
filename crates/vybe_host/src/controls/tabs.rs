use super::{ControlDef, Props, base_css};

pub static TABCONTROL_DEF: ControlDef = ControlDef {
    tag: "div",
    inner_tag: None,
    props: &["SelectedIndex", "Visible", "BackColor"],
    events: &["SelectedIndexChanged"],
    default_size: (200, 100),
    css_fn: tabcontrol_css,
    container: true,
    input_type: None,
    extra_attrs: &[],
};

pub static TABPAGE_DEF: ControlDef = ControlDef {
    tag: "div",
    inner_tag: None,
    props: &["Text", "Visible", "BackColor"],
    events: &[],
    default_size: (200, 100),
    css_fn: tabpage_css,
    container: true,
    input_type: None,
    extra_attrs: &[],
};

fn tabcontrol_css(props: &Props) -> String {
    let mut s = String::from("position: relative; border: 1px solid #ccc; ");
    s.push_str(&base_css(props));
    s
}

fn tabpage_css(props: &Props) -> String {
    let mut s = String::from("position: relative; padding: 8px; ");
    s.push_str(&base_css(props));
    s
}
