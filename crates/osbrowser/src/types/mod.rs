//! The document and its data types, as webcore exposes them.
//!
//! Same names, fields and signatures as `webcore::types`, so code written
//! against one browser compiles against the other. The tree itself lives in the
//! system browser; [`Document`] is this process's handle on it.

mod document;

pub use document::{Document, DocumentRoot};

/// The DOM's own distinction, and the only thing it changes here is whether
/// names FOLD: HTML is ASCII-case-insensitive for tag and attribute names, XML
/// is not.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum DocumentKind {
    Html,
    Xml,
}

/// Form interaction event — fired by the engine, handled by the host.
#[derive(Debug, Clone)]
pub struct FormEvent {
    /// Element tag (e.g. "input", "select", "textarea")
    pub tag: String,
    /// Element id attribute (empty if none)
    pub id: String,
    /// Element name attribute (empty if none)
    pub name: String,
    /// Event kind
    pub kind: FormEventKind,
    /// Stable node_id of the element.
    pub element: u32,
}

#[derive(Debug, Clone)]
pub enum FormEventKind {
    /// Text input value changed (new value)
    Input(String),
    /// Value committed (e.g. Enter in text field, option selected)
    Change(String),
    /// Checkbox/radio toggled (new checked state)
    Toggle(bool),
    /// Button clicked (value attribute)
    Click(String),
    /// Form submitted (form element's action URL)
    Submit(String),
    /// Focus gained
    Focus,
    /// Focus lost
    Blur,
}

pub type FormEventCallback = Box<dyn FnMut(&FormEvent) + Send>;

#[derive(Clone, Copy, Debug, Default, PartialEq, serde::Deserialize)]
pub struct Rect {
    pub x: f32,
    pub y: f32,
    pub w: f32,
    pub h: f32,
}

#[derive(Clone, Debug, Default, PartialEq)]
pub struct PendingImagePoll {
    pub loaded_any: bool,
    pub needs_relayout: bool,
    pub paint_rects: Vec<Rect>,
}
