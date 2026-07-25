use super::{ControlDef, Props, base_css};

pub static DEF: ControlDef = ControlDef {
    tag: "div",
    inner_tag: None,
    props: &["Text", "Visible", "BackColor", "BorderStyle", "AutoScroll"],
    events: &["Click", "Paint"],
    default_size: (200, 100),
    css_fn: css,
    container: true,
    input_type: None,
    extra_attrs: &[],
};

pub static SPLIT_DEF: ControlDef = ControlDef {
    tag: "div",
    inner_tag: None,
    props: &["Orientation", "SplitterDistance", "Visible"],
    events: &["SplitterMoved"],
    default_size: (200, 100),
    css_fn: split_css,
    container: true,
    input_type: None,
    extra_attrs: &[],
};

pub static FLOW_DEF: ControlDef = ControlDef {
    tag: "div",
    inner_tag: None,
    props: &["FlowDirection", "WrapContents", "Visible", "BackColor"],
    events: &[],
    default_size: (200, 100),
    css_fn: flow_css,
    container: true,
    input_type: None,
    extra_attrs: &[],
};

pub static TABLE_DEF: ControlDef = ControlDef {
    tag: "div",
    inner_tag: None,
    props: &["ColumnCount", "RowCount", "Visible", "BackColor"],
    events: &[],
    default_size: (200, 100),
    css_fn: table_css,
    container: true,
    input_type: None,
    extra_attrs: &[],
};

fn css(props: &Props) -> String {
    let mut s = String::from("position: relative; overflow: hidden; ");
    match props.get("BorderStyle").map(|v| v.as_str()) {
        Some("FixedSingle") => s.push_str("border: 1px solid #999; "),
        Some("Fixed3D") => s.push_str("border: 2px inset #ccc; "),
        _ => {}
    }
    if props
        .get("AutoScroll")
        .map(|v| v == "True")
        .unwrap_or(false)
    {
        s.push_str("overflow: auto; ");
    }
    s.push_str(&base_css(props));
    s
}

fn split_css(props: &Props) -> String {
    let mut s = String::from("display: flex; ");
    if props
        .get("Orientation")
        .map(|v| v == "Horizontal")
        .unwrap_or(false)
    {
        s.push_str("flex-direction: column; ");
    }
    s.push_str(&base_css(props));
    s
}

fn flow_css(props: &Props) -> String {
    let mut s = String::from("display: flex; flex-wrap: wrap; gap: 4px; ");
    if props
        .get("FlowDirection")
        .map(|v| v == "TopDown")
        .unwrap_or(false)
    {
        s.push_str("flex-direction: column; ");
    }
    s.push_str(&base_css(props));
    s
}

fn table_css(props: &Props) -> String {
    let cols = props
        .get("ColumnCount")
        .and_then(|v| v.parse::<u32>().ok())
        .unwrap_or(2);
    let mut s = format!(
        "display: grid; grid-template-columns: repeat({}, 1fr); gap: 2px; ",
        cols
    );
    s.push_str(&base_css(props));
    s
}
