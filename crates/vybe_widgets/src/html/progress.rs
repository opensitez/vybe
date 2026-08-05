use super::{ControlDef, Props, base_css};

pub static PROGRESS_DEF: ControlDef = ControlDef {
    tag: "progress",
    inner_tag: None,
    props: &["Value", "Minimum", "Maximum", "Visible", "Style"],
    events: &[],
    default_size: (100, 23),
    css_fn: progress_css,
    container: false,
    input_type: None,
    extra_attrs: &[],
};

pub static TRACKBAR_DEF: ControlDef = ControlDef {
    tag: "input",
    inner_tag: None,
    props: &["Value", "Minimum", "Maximum", "TickFrequency", "Visible"],
    events: &["ValueChanged", "Scroll"],
    default_size: (104, 45),
    css_fn: trackbar_css,
    container: false,
    input_type: Some("range"),
    extra_attrs: &[],
};

pub static SCROLLBAR_DEF: ControlDef = ControlDef {
    tag: "input",
    inner_tag: None,
    props: &["Value", "Minimum", "Maximum", "Visible"],
    events: &["ValueChanged", "Scroll"],
    default_size: (80, 17),
    css_fn: trackbar_css,
    container: false,
    input_type: Some("range"),
    extra_attrs: &[],
};

fn progress_css(props: &Props) -> String {
    let mut s = String::from("appearance: auto; ");
    s.push_str(&base_css(props));
    s
}

fn trackbar_css(props: &Props) -> String {
    let mut s = String::from("appearance: auto; cursor: pointer; ");
    s.push_str(&base_css(props));
    s
}
