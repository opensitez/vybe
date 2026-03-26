use super::{ControlDef, Props, base_css};

pub static DATAGRID_DEF: ControlDef = ControlDef {
    tag: "table",
    inner_tag: None,
    props: &["Columns", "Rows", "Visible", "BackColor", "ReadOnly", "AllowUserToAddRows"],
    events: &["CellClick", "CellValueChanged", "SelectionChanged", "RowEnter"],
    default_size: (240, 150),
    css_fn: grid_css,
    container: false,
    input_type: None,
    extra_attrs: &[],
};

pub static LISTVIEW_DEF: ControlDef = ControlDef {
    tag: "div",
    inner_tag: None,
    props: &["Items", "Columns", "View", "Visible", "BackColor"],
    events: &["SelectedIndexChanged", "Click", "DoubleClick", "ItemActivate"],
    default_size: (120, 97),
    css_fn: listview_css,
    container: false,
    input_type: None,
    extra_attrs: &[],
};

fn grid_css(props: &Props) -> String {
    let mut s = String::from("border-collapse: collapse; border: 1px solid #ccc; width: 100%; height: 100%; overflow: auto; display: block; ");
    s.push_str(&base_css(props));
    s
}

fn listview_css(props: &Props) -> String {
    let mut s = String::from("border: 1px solid #ccc; overflow: auto; ");
    s.push_str(&base_css(props));
    s
}
