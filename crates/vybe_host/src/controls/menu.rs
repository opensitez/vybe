use super::{ControlDef, Props, base_css};

pub static MENUSTRIP_DEF: ControlDef = ControlDef {
    tag: "nav",
    inner_tag: None,
    props: &["Items", "Visible", "BackColor"],
    events: &["ItemClicked"],
    default_size: (200, 24),
    css_fn: menu_css,
    container: true,
    input_type: None,
    extra_attrs: &[],
};

pub static CONTEXTMENU_DEF: ControlDef = ControlDef {
    tag: "div",
    inner_tag: None,
    props: &["Items", "Visible"],
    events: &["ItemClicked", "Opening"],
    default_size: (120, 0),
    css_fn: context_css,
    container: true,
    input_type: None,
    extra_attrs: &[],
};

pub static TOOLSTRIP_DEF: ControlDef = ControlDef {
    tag: "div",
    inner_tag: None,
    props: &["Items", "Visible", "BackColor"],
    events: &["ItemClicked"],
    default_size: (200, 25),
    css_fn: toolbar_css,
    container: true,
    input_type: None,
    extra_attrs: &[],
};

pub static STATUSSTRIP_DEF: ControlDef = ControlDef {
    tag: "div",
    inner_tag: None,
    props: &["Items", "Text", "Visible", "BackColor"],
    events: &[],
    default_size: (200, 22),
    css_fn: status_css,
    container: true,
    input_type: None,
    extra_attrs: &[],
};

fn menu_css(props: &Props) -> String {
    let mut s = String::from(
        "display: flex; gap: 0; background: #f0f0f0; border-bottom: 1px solid #ccc; padding: 2px 4px; ",
    );
    s.push_str(&base_css(props));
    s
}

fn context_css(props: &Props) -> String {
    let mut s = String::from(
        "position: absolute; background: white; border: 1px solid #ccc; box-shadow: 2px 2px 4px rgba(0,0,0,0.2); z-index: 1000; padding: 4px 0; ",
    );
    s.push_str(&base_css(props));
    s
}

fn toolbar_css(props: &Props) -> String {
    let mut s = String::from(
        "display: flex; gap: 2px; align-items: center; background: #f0f0f0; border-bottom: 1px solid #ccc; padding: 2px 4px; ",
    );
    s.push_str(&base_css(props));
    s
}

fn status_css(props: &Props) -> String {
    let mut s = String::from(
        "display: flex; align-items: center; background: #f0f0f0; border-top: 1px solid #ccc; padding: 2px 8px; font-size: 12px; ",
    );
    s.push_str(&base_css(props));
    s
}
