use super::{ControlDef, Props, base_css};

pub static DEF: ControlDef = ControlDef {
    tag: "select",
    inner_tag: None,
    props: &["Items", "SelectedIndex", "SelectedItem", "Enabled", "Visible", "SelectionMode"],
    events: &["SelectedIndexChanged", "Click", "DoubleClick"],
    default_size: (120, 95),
    css_fn: css,
    container: false,
    input_type: None,
    extra_attrs: &[("multiple", ""), ("size", "5")],
};

fn css(props: &Props) -> String {
    let mut s = String::from("padding: 2px; border: 1px solid #999; box-sizing: border-box; overflow-y: auto; ");
    s.push_str(&base_css(props));
    s
}
