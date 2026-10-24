//! The DOM, as webcore exposes it: event types, events and the document
//! registry. Same paths, types and signatures as `webcore::dom`.

pub mod events;
pub mod registry;

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum HtmlEventType {
    // Mouse
    Click,
    DblClick,
    MouseDown,
    MouseUp,
    MouseMove,
    /// Fires when cursor enters element or any descendant (bubbles). Mirror of CSS :hover.
    MouseOver,
    /// Fires when cursor leaves element or any descendant (bubbles).
    MouseOut,
    /// Fires when cursor enters element boundary (does not bubble).
    MouseEnter,
    /// Fires when cursor leaves element boundary (does not bubble).
    MouseLeave,
    ContextMenu,
    // Pointer (unified mouse/touch/stylus — fired alongside mouse events on desktop)
    PointerDown,
    PointerUp,
    PointerMove,
    PointerCancel,
    PointerEnter,
    PointerLeave,
    PointerOver,
    PointerOut,
    // Drag
    DragStart,
    Drag,
    DragEnter,
    DragOver,
    DragLeave,
    Drop,
    DragEnd,
    // Keyboard
    KeyDown,
    KeyUp,
    KeyPress,
    // Wheel / scroll
    /// Raw wheel input from mouse or trackpad (fires before Scroll).
    Wheel,
    Scroll,
    // Content / form / focus
    Input,
    Change,
    /// Fires on the focused element (does not bubble).
    Focus,
    /// Fires on the focused element when it loses focus (does not bubble).
    Blur,
    /// Bubbling version of Focus.
    FocusIn,
    /// Bubbling version of Blur.
    FocusOut,
    SelectionChange,
    // Document / window lifecycle
    /// Fires on the document root after parsing and layout complete.
    DOMContentLoaded,
    /// Fires on the document root when the document is about to be replaced or the window closed.
    Unload,
    /// Fires on the document root when the viewport is resized.
    Resize,
}

impl HtmlEventType {
    /// DOM event type string (e.g. "click", "mousedown", "keydown").
    pub fn as_str(&self) -> &'static str {
        match self {
            Self::Click => "click",
            Self::DblClick => "dblclick",
            Self::MouseDown => "mousedown",
            Self::MouseUp => "mouseup",
            Self::MouseMove => "mousemove",
            Self::MouseOver => "mouseover",
            Self::MouseOut => "mouseout",
            Self::MouseEnter => "mouseenter",
            Self::MouseLeave => "mouseleave",
            Self::ContextMenu => "contextmenu",
            Self::PointerDown => "pointerdown",
            Self::PointerUp => "pointerup",
            Self::PointerMove => "pointermove",
            Self::PointerCancel => "pointercancel",
            Self::PointerEnter => "pointerenter",
            Self::PointerLeave => "pointerleave",
            Self::PointerOver => "pointerover",
            Self::PointerOut => "pointerout",
            Self::DragStart => "dragstart",
            Self::Drag => "drag",
            Self::DragEnter => "dragenter",
            Self::DragOver => "dragover",
            Self::DragLeave => "dragleave",
            Self::Drop => "drop",
            Self::DragEnd => "dragend",
            Self::KeyDown => "keydown",
            Self::KeyUp => "keyup",
            Self::KeyPress => "keypress",
            Self::Wheel => "wheel",
            Self::Scroll => "scroll",
            Self::Input => "input",
            Self::Change => "change",
            Self::Focus => "focus",
            Self::Blur => "blur",
            Self::FocusIn => "focusin",
            Self::FocusOut => "focusout",
            Self::SelectionChange => "selectionchange",
            Self::DOMContentLoaded => "DOMContentLoaded",
            Self::Unload => "unload",
            Self::Resize => "resize",
        }
    }
}

pub struct HtmlEvent {
    pub event_type: HtmlEventType,
    /// Stable node_id of the deepest box hit.
    pub target: u32,
    /// Stable node_id of the current listener's element.
    pub current_target: u32,
    /// Position in window coordinates.
    pub client_pos: (f32, f32),
    /// Position in document coordinates.
    pub doc_pos: (f32, f32),
    /// Mouse button: 0=left, 1=middle, 2=right.
    pub button: u8,
    pub key_code: u32,
    pub char_code: Option<char>,
    pub ctrl_key: bool,
    pub shift_key: bool,
    pub alt_key: bool,
    pub meta_key: bool,
    /// Wheel scroll delta in logical pixels (positive = scroll down/right).
    pub delta_x: f32,
    pub delta_y: f32,
    /// Source node_id for drag events.
    pub drag_source: u32,
    /// Related target node_id (e.g. focus/blur counterpart).
    pub related_target: u32,
    pub default_prevented: bool,
    pub propagation_stopped: bool,
}

impl HtmlEvent {
    pub fn new(event_type: HtmlEventType) -> Self {
        Self {
            event_type,
            target: 0,
            current_target: 0,
            client_pos: (0.0, 0.0),
            doc_pos: (0.0, 0.0),
            button: 0,
            key_code: 0,
            char_code: None,
            ctrl_key: false,
            shift_key: false,
            alt_key: false,
            meta_key: false,
            delta_x: 0.0,
            delta_y: 0.0,
            drag_source: 0,
            related_target: 0,
            default_prevented: false,
            propagation_stopped: false,
        }
    }

    pub fn stop_propagation(&mut self) {
        self.propagation_stopped = true;
    }
    pub fn prevent_default(&mut self) {
        self.default_prevented = true;
    }
}
