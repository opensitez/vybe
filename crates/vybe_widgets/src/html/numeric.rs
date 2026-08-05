use super::{ControlDef, Props, base_css};

pub static DEF: ControlDef = ControlDef {
    tag: "input",
    inner_tag: None,
    props: &[
        "Value",
        "Minimum",
        "Maximum",
        "DecimalPlaces",
        "Increment",
        "Enabled",
        "Visible",
    ],
    events: &["ValueChanged"],
    default_size: (120, 23),
    css_fn: css,
    container: false,
    input_type: Some("number"),
    extra_attrs: &[],
};

fn css(props: &Props) -> String {
    let mut s = String::from("padding: 2px 4px; border: 1px solid #999; box-sizing: border-box; ");
    s.push_str(&base_css(props));
    s
}
