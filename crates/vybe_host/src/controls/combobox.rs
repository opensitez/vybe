use super::{ControlDef, Props, base_css};

pub static DEF: ControlDef = ControlDef {
    tag: "select",
    inner_tag: None,
    props: &[
        "Text",
        "Items",
        "SelectedIndex",
        "SelectedItem",
        "Enabled",
        "Visible",
        "DropDownStyle",
    ],
    events: &["SelectedIndexChanged", "TextChanged", "DropDown"],
    default_size: (120, 23),
    css_fn: css,
    container: false,
    input_type: None,
    extra_attrs: &[],
};

fn css(props: &Props) -> String {
    let mut s = String::from("padding: 2px; border: 1px solid #999; box-sizing: border-box; ");
    s.push_str(&base_css(props));
    s
}
