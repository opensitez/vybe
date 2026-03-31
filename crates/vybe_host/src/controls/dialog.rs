use super::ControlDef;

// Dialogs are non-visual — they don't render, they produce side effects.
// These defs are mainly for documentation and property/event listing.

pub static OPEN_FILE_DEF: ControlDef = ControlDef {
    tag: "div",
    inner_tag: None,
    props: &["Filter", "Title", "FileName", "Multiselect", "InitialDirectory"],
    events: &["FileOk"],
    default_size: (0, 0),
    css_fn: |_| String::new(),
    container: false,
    input_type: None,
    extra_attrs: &[],
};

pub static SAVE_FILE_DEF: ControlDef = ControlDef {
    tag: "div",
    inner_tag: None,
    props: &["Filter", "Title", "FileName", "DefaultExt", "InitialDirectory"],
    events: &["FileOk"],
    default_size: (0, 0),
    css_fn: |_| String::new(),
    container: false,
    input_type: None,
    extra_attrs: &[],
};
