use super::{ControlDef, Props, base_css};

pub static DATETIME_DEF: ControlDef = ControlDef {
    tag: "input",
    inner_tag: None,
    props: &["Value", "Format", "MinDate", "MaxDate", "Visible", "Enabled"],
    events: &["ValueChanged"],
    default_size: (200, 23),
    css_fn: datetime_css,
    container: false,
    input_type: Some("date"),
    extra_attrs: &[],
};

pub static CALENDAR_DEF: ControlDef = ControlDef {
    tag: "input",
    inner_tag: None,
    props: &["SelectionStart", "SelectionEnd", "MinDate", "MaxDate", "Visible"],
    events: &["DateChanged", "DateSelected"],
    default_size: (199, 162),
    css_fn: calendar_css,
    container: false,
    input_type: Some("date"),
    extra_attrs: &[],
};

fn datetime_css(props: &Props) -> String {
    let mut s = String::from("padding: 2px 4px; border: 1px solid #999; box-sizing: border-box; ");
    s.push_str(&base_css(props));
    s
}

fn calendar_css(props: &Props) -> String {
    let mut s = String::from("border: 1px solid #999; ");
    s.push_str(&base_css(props));
    s
}
