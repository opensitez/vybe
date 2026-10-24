//! DOM events, as webcore declares them.
//!
//! The same types and methods as `webcore::dom::events`, so a caller written
//! against one browser compiles against the other. Listeners registered through
//! `Document::add_event_listener` live in this process; the system browser
//! reports each event it delivers and the handler runs here.

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum EventPhase {
    None,
    Capture,
    Target,
    Bubble,
}

/// DOM event — flows through capture → target → bubble phases.
#[derive(Debug)]
pub struct DomEvent {
    pub event_type: String,
    pub target: u32,         // node_id of the deepest hit element
    pub current_target: u32, // node_id of the element currently handling
    pub phase: EventPhase,
    /// Mouse/pointer position in document coordinates.
    pub client_x: f32,
    pub client_y: f32,
    /// Mouse button (0=left, 1=middle, 2=right).
    pub button: u8,
    /// Keyboard key code.
    pub key_code: u32,
    pub key: String,
    pub char_code: Option<char>,
    /// Modifier keys.
    pub ctrl_key: bool,
    pub shift_key: bool,
    pub alt_key: bool,
    pub meta_key: bool,
    /// Wheel delta.
    pub delta_x: f32,
    pub delta_y: f32,
    /// Related target (e.g. for mouseenter/leave, the element being entered/left).
    pub related_target: u32,
    /// `Event.bubbles` — does this event run the bubble phase at all?
    ///
    /// Not every event does. `load`, `focus`, `mouseenter` and friends fire on
    /// the target and stop; dispatching them up the tree calls listeners a
    /// browser never would. Set from `event_defaults` by `DomEvent::new`.
    pub bubbles: bool,
    /// `Event.cancelable` — may `preventDefault()` do anything?
    pub cancelable: bool,
    /// `Event.composed` — does this event cross a shadow boundary?
    pub composed: bool,
    /// `Event.isTrusted` — true when the user agent generated it, false when a
    /// script did (DOM §2.2). `dispatch_event` from the API sets it false.
    pub is_trusted: bool,
    /// `Event.timeStamp` — milliseconds since the document began.
    pub time_stamp: f64,
    /// `Event.composedPath()` — the propagation path, target-first. Filled by
    /// dispatch; empty before and after.
    pub composed_path: Vec<u32>,

    /// `ToggleEvent.oldState` / `.newState` (HTML §6.12) — `"open"` or
    /// `"closed"`. Empty for every other event type, like the mouse and
    /// keyboard fields above, which this struct already carries for all of
    /// them rather than splitting into one type per interface.
    pub old_state: String,
    pub new_state: String,
    /// `MediaQueryListEvent.media` / `.matches` (cssom-view-1 §4.2).
    pub media: String,
    pub matches: bool,
    /// HashChangeEvent URL payload.
    pub old_url: String,
    pub new_url: String,
    /// `CommandEvent.command` and its invoking button.
    pub command: String,
    pub source: u32,
    /// CSS AnimationEvent / TransitionEvent payload. Elapsed time is in seconds
    /// of active time, excluding the delay and any paused interval.
    pub elapsed_time: f64,
    pub animation_name: String,
    pub property_name: String,
    pub pseudo_element: String,
    /// `CustomEvent.detail`.
    detail: String,
    /// `UIEvent.detail` — click count.
    ui_detail: i32,
    /// `MouseEvent.buttons` bitmask: 1 left, 2 right, 4 middle.
    buttons: u8,
    /// `MouseEvent.screenX/Y`.
    screen_x: f32,
    screen_y: f32,
    /// Page scroll at dispatch, for `pageX`/`pageY`.
    scroll_x: f32,
    scroll_y: f32,
    /// The target's padding-box origin, for `offsetX`/`offsetY`.
    target_origin: (f32, f32),
    /// `KeyboardEvent.code` — the physical key.
    code: String,
    /// `KeyboardEvent.location`.
    location: u32,
    /// `KeyboardEvent.repeat`.
    repeat: bool,
    /// `KeyboardEvent.isComposing`.
    is_composing: bool,
    /// `WheelEvent.deltaZ` and `deltaMode`.
    delta_z: f32,
    delta_mode: u32,
    /// `InputEvent.data` and `inputType`.
    data: Option<String>,
    input_type: String,
    clipboard_data: Option<ClipboardData>,
    /// Call to prevent default browser behavior.
    prevented: bool,
    /// Call to stop event from reaching further listeners.
    stopped: bool,
    /// Call to stop event from reaching listeners on the same element.
    immediate_stopped: bool,
    /// DOM §2.9's "dispatch flag". `dispatchEvent` on an event that is already
    /// being dispatched must fail rather than re-enter — a handler calling
    /// `dispatchEvent(sameEvent)` would otherwise recurse forever.
    dispatching: bool,
}

/// `bubbles`, `cancelable` and `composed` for an event type (HTML §4.11 and the
/// UI Events spec).
///
/// These are per-TYPE facts, not per-dispatch choices: `click` bubbles and is
/// cancelable, `load` is neither, `focus` is neither but `focusin` bubbles.
/// Getting them wrong is not cosmetic — a non-bubbling event dispatched up the
/// tree runs listeners a browser never runs.
pub fn event_defaults(event_type: &str) -> (bool, bool, bool) {
    match event_type {
        // Mouse and pointer — bubble, cancelable, composed.
        "click" | "auxclick" | "dblclick" | "mousedown" | "mouseup" | "mousemove" | "mouseover"
        | "mouseout" | "contextmenu" | "wheel" | "pointerdown" | "pointerup" | "pointermove"
        | "pointerover" | "pointerout" | "keydown" | "keyup" | "keypress" | "beforeinput"
        | "compositionstart" | "compositionupdate" | "compositionend" | "cut" | "copy"
        | "paste" | "dragstart" | "drag" | "dragenter" | "dragover" | "dragleave" | "drop"
        | "submit" | "reset" | "beforetoggle" => (true, true, true),
        // Bubble, not cancelable.
        "input" | "change" | "select" | "focusin" | "focusout" | "dragend" | "scrollend"
        | "pointercancel" | "formdata" | "slotchange" => (true, false, true),
        // Fire on the target only, not cancelable.
        "load" | "unload" | "error" | "abort" | "focus" | "blur" | "mouseenter" | "mouseleave"
        | "pointerenter" | "pointerleave" | "scroll" | "resize" | "toggle"
        | "cancel" | "close" | "canplay" | "canplaythrough" | "durationchange" | "emptied"
        | "ended" | "loadeddata" | "loadedmetadata" | "loadstart" | "pause" | "play"
        | "playing" | "progress" | "ratechange" | "seeked" | "seeking" | "stalled" | "suspend"
        | "timeupdate" | "volumechange" | "waiting" => (false, false, false),
        // These do not bubble, but their default behavior can be canceled.
        "beforeunload" | "invalid" => (false, true, false),
        // Window/document lifecycle — bubble, not cancelable.
        "hashchange" | "popstate" | "pagehide" | "pageshow" | "offline" | "online"
        | "languagechange" | "storage" | "message" | "messageerror" | "afterprint"
        | "beforeprint" | "visibilitychange" | "readystatechange" | "DOMContentLoaded" => {
            (true, false, false)
        }
        // CSS animation/transition events bubble but are not cancelable or composed.
        "animationstart" | "animationend" | "animationiteration" | "animationcancel"
        | "transitionstart" | "transitionend" | "transitionrun" | "transitioncancel" => {
            (true, false, false)
        }
        // Anything else, including custom events, defaults to the DOM's own
        // defaults: an event created with no options bubbles nowhere.
        _ => (false, false, false),
    }
}

impl DomEvent {
    /// An event the USER AGENT fires: flags come from the per-type table and
    /// `isTrusted` is true. For a script-constructed event use
    /// [`DomEvent::new_script`] or [`DomEvent::new_with_flags`] — those default
    /// to not bubbling, which is what `new Event(type)` does in a browser.
    pub fn new(event_type: impl Into<String>, target: u32) -> Self {
        let ty: String = event_type.into();
        let (bubbles, cancelable, composed) = event_defaults(&ty);
        Self {
            event_type: ty,
            target,
            current_target: 0,
            phase: EventPhase::None,
            client_x: 0.0,
            client_y: 0.0,
            button: 0,
            key_code: 0,
            key: String::new(),
            char_code: None,
            ctrl_key: false,
            shift_key: false,
            alt_key: false,
            meta_key: false,
            delta_x: 0.0,
            delta_y: 0.0,
            related_target: 0,
            bubbles,
            cancelable,
            composed,
            is_trusted: true,
            time_stamp: time_origin_elapsed_ms(),
            composed_path: Vec::new(),
            old_state: String::new(),
            new_state: String::new(),
            media: String::new(),
            matches: false,
            old_url: String::new(),
            new_url: String::new(),
            command: String::new(),
            source: 0,
            elapsed_time: 0.0,
            animation_name: String::new(),
            property_name: String::new(),
            pseudo_element: String::new(),
            detail: String::new(),
            ui_detail: 0,
            buttons: 0,
            screen_x: 0.0,
            screen_y: 0.0,
            scroll_x: 0.0,
            scroll_y: 0.0,
            target_origin: (0.0, 0.0),
            code: String::new(),
            location: 0,
            repeat: false,
            is_composing: false,
            delta_z: 0.0,
            delta_mode: 0,
            data: None,
            input_type: String::new(),
            clipboard_data: None,
            prevented: false,
            stopped: false,
            immediate_stopped: false,
            dispatching: false,
        }
    }

    /// `new Event(type)` — the script constructor with no options.
    ///
    /// **All three flags default to FALSE, whatever the type is.** The per-type
    /// table only describes events the USER AGENT fires; a script-constructed
    /// `new Event("click")` does not bubble unless the caller asks for it.
    /// Confirmed against Chrome. Use [`DomEvent::new`] for a UA-fired event and
    /// this for anything a script makes.
    pub fn new_script(event_type: impl Into<String>, target: u32) -> Self {
        Self::new_with_flags(event_type, target, false, false, false)
    }

    /// `new Event(type, { bubbles, cancelable, composed })` — a script-created
    /// event, which is never trusted and takes its flags from the caller
    /// rather than from the type table.
    pub fn new_with_flags(
        event_type: impl Into<String>,
        target: u32,
        bubbles: bool,
        cancelable: bool,
        composed: bool,
    ) -> Self {
        let mut e = Self::new(event_type, target);
        e.bubbles = bubbles;
        e.cancelable = cancelable;
        e.composed = composed;
        e.is_trusted = false;
        e
    }

    /// `preventDefault()` — only does anything on a CANCELABLE event
    /// (DOM §2.3). Calling it on `load` or `mouseenter` is a no-op in a
    /// browser, and treating it as one here keeps `defaultPrevented` honest.
    pub fn prevent_default(&mut self) {
        if self.cancelable {
            self.prevented = true;
        }
    }
    pub fn stop_propagation(&mut self) {
        self.stopped = true;
    }
    pub fn stop_immediate_propagation(&mut self) {
        self.stopped = true;
        self.immediate_stopped = true;
    }
    pub fn default_prevented(&self) -> bool {
        self.prevented
    }
    pub fn propagation_stopped(&self) -> bool {
        self.stopped
    }

    /// `Event.bubbles`.
    pub fn bubbles(&self) -> bool {
        self.bubbles
    }
    /// `Event.cancelable`.
    pub fn cancelable(&self) -> bool {
        self.cancelable
    }
    /// `Event.composed`.
    pub fn composed(&self) -> bool {
        self.composed
    }
    /// `Event.isTrusted` — true only for events the user agent generated.
    pub fn is_trusted(&self) -> bool {
        self.is_trusted
    }
    /// `Event.timeStamp` — milliseconds since the engine's time origin.
    pub fn time_stamp(&self) -> f64 {
        self.time_stamp
    }

    /// Is this event mid-dispatch? `dispatchEvent` refuses to re-enter it.
    pub fn is_dispatching(&self) -> bool {
        self.dispatching
    }

    /// `Event.eventPhase` — the numeric constant DOM §2.2 defines.
    pub fn event_phase(&self) -> u16 {
        match self.phase {
            EventPhase::None => 0,
            EventPhase::Capture => 1,
            EventPhase::Target => 2,
            EventPhase::Bubble => 3,
        }
    }

    /// `Event.composedPath()` — target first, root last.
    pub fn composed_path(&self) -> &[u32] {
        &self.composed_path
    }

    /// `Event.srcElement` — the legacy alias for `target`, still shipped.
    pub fn src_element(&self) -> u32 {
        self.target
    }

    /// `Event.cancelBubble` — legacy alias for "propagation stopped".
    pub fn cancel_bubble(&self) -> bool {
        self.stopped
    }
    pub fn set_cancel_bubble(&mut self, v: bool) {
        if v {
            self.stopped = true;
        }
    }

    /// `Event.returnValue` — legacy inverse of `defaultPrevented`.
    pub fn return_value(&self) -> bool {
        !self.prevented
    }
    pub fn set_return_value(&mut self, v: bool) {
        if !v {
            self.prevent_default();
        }
    }

    /// `Event.initEvent(type, bubbles, cancelable)` — legacy initialiser. Does
    /// nothing once the event is being dispatched, per DOM §2.2.
    pub fn init_event(&mut self, event_type: &str, bubbles: bool, cancelable: bool) {
        if self.phase != EventPhase::None {
            return;
        }
        self.event_type = event_type.to_string();
        self.bubbles = bubbles;
        self.cancelable = cancelable;
        self.prevented = false;
        self.stopped = false;
        self.immediate_stopped = false;
        self.is_trusted = false;
    }
}

impl DomEvent {
    // ── UIEvent ──

    /// `UIEvent.detail` — for a click, the click count (1 single, 2 double).
    /// Zero for events that do not count.
    pub fn ui_detail(&self) -> i32 {
        self.ui_detail
    }
    pub fn set_ui_detail(&mut self, n: i32) {
        self.ui_detail = n;
    }

    // ── MouseEvent ──

    /// `MouseEvent.clientX` / `clientY` — viewport coordinates.
    pub fn client_x(&self) -> f32 {
        self.client_x
    }
    pub fn client_y(&self) -> f32 {
        self.client_y
    }

    /// `MouseEvent.pageX` / `pageY` — document coordinates: the viewport
    /// position plus how far the page is scrolled.
    pub fn page_x(&self) -> f32 {
        self.client_x + self.scroll_x
    }
    pub fn page_y(&self) -> f32 {
        self.client_y + self.scroll_y
    }

    /// `MouseEvent.screenX` / `screenY`.
    pub fn screen_x(&self) -> f32 {
        self.screen_x
    }
    pub fn screen_y(&self) -> f32 {
        self.screen_y
    }

    /// `MouseEvent.offsetX` / `offsetY` — relative to the target's padding box.
    pub fn offset_x(&self) -> f32 {
        self.client_x - self.target_origin.0
    }
    pub fn offset_y(&self) -> f32 {
        self.client_y - self.target_origin.1
    }

    /// `MouseEvent.button` — which button CHANGED state: 0 left, 1 middle,
    /// 2 right. Meaningless on `mousemove`, where nothing changed.
    pub fn button(&self) -> u8 {
        self.button
    }

    /// `MouseEvent.buttons` — a BITMASK of which buttons are currently held:
    /// 1 left, 2 right, 4 middle. A different question from `button`, and the
    /// one to ask during a drag. Note bit 2 is RIGHT and bit 4 is MIDDLE —
    /// the order is not the same as `button`'s numbering.
    pub fn buttons(&self) -> u8 {
        self.buttons
    }
    pub fn set_buttons(&mut self, mask: u8) {
        self.buttons = mask;
    }

    /// `MouseEvent.relatedTarget` — 0 when there is none.
    pub fn related_target(&self) -> u32 {
        self.related_target
    }

    // ── Modifier keys, shared by mouse and keyboard events ──

    pub fn ctrl_key(&self) -> bool {
        self.ctrl_key
    }
    pub fn shift_key(&self) -> bool {
        self.shift_key
    }
    pub fn alt_key(&self) -> bool {
        self.alt_key
    }
    pub fn meta_key(&self) -> bool {
        self.meta_key
    }

    /// `getModifierState(key)` — UI Events §5.2. Recognises the modifier key
    /// VALUES the spec lists, not arbitrary names; an unknown name is false.
    pub fn get_modifier_state(&self, key: &str) -> bool {
        match key {
            "Control" => self.ctrl_key,
            "Shift" => self.shift_key,
            "Alt" => self.alt_key,
            "Meta" => self.meta_key,
            "AltGraph" | "CapsLock" | "Fn" | "FnLock" | "Hyper" | "NumLock" | "ScrollLock"
            | "Super" | "Symbol" | "SymbolLock" => false,
            _ => false,
        }
    }

    // ── KeyboardEvent ──

    /// `KeyboardEvent.key` — the character or named key produced, e.g. `a`,
    /// `A`, `Enter`, `ArrowLeft`.
    pub fn key(&self) -> &str {
        &self.key
    }

    /// `KeyboardEvent.code` — the PHYSICAL key, e.g. `KeyA`, independent of
    /// layout and of shift. `key` and `code` differ on any non-US layout.
    pub fn code(&self) -> &str {
        &self.code
    }
    pub fn set_code(&mut self, code: impl Into<String>) {
        self.code = code.into();
    }

    /// `KeyboardEvent.location` — 0 standard, 1 left, 2 right, 3 numpad.
    pub fn location(&self) -> u32 {
        self.location
    }
    pub fn set_location(&mut self, loc: u32) {
        self.location = loc;
    }

    /// `KeyboardEvent.repeat` — held down and auto-repeating.
    pub fn repeat(&self) -> bool {
        self.repeat
    }
    pub fn set_repeat(&mut self, r: bool) {
        self.repeat = r;
    }

    /// `KeyboardEvent.isComposing` — inside an IME composition session.
    pub fn is_composing(&self) -> bool {
        self.is_composing
    }
    pub fn set_is_composing(&mut self, c: bool) {
        self.is_composing = c;
    }

    // ── WheelEvent ──

    /// `WheelEvent.deltaX` / `deltaY` / `deltaZ`.
    pub fn delta_x(&self) -> f32 {
        self.delta_x
    }
    pub fn delta_y(&self) -> f32 {
        self.delta_y
    }
    pub fn delta_z(&self) -> f32 {
        self.delta_z
    }

    /// `WheelEvent.deltaMode` — 0 pixels, 1 lines, 2 pages.
    pub fn delta_mode(&self) -> u32 {
        self.delta_mode
    }
    pub fn set_delta_mode(&mut self, m: u32) {
        self.delta_mode = m;
    }

    // ── InputEvent ──

    /// `InputEvent.data` — the characters inserted, if any.
    pub fn data(&self) -> Option<&str> {
        self.data.as_deref()
    }
    pub fn set_data(&mut self, d: Option<String>) {
        self.data = d;
    }

    /// `InputEvent.inputType` — e.g. `insertText`, `deleteContentBackward`.
    pub fn input_type(&self) -> &str {
        &self.input_type
    }
    pub fn set_input_type(&mut self, t: impl Into<String>) {
        self.input_type = t.into();
    }

    pub fn clipboard_data(&self) -> Option<&ClipboardData> {
        self.clipboard_data.as_ref()
    }

    pub fn clipboard_data_mut(&mut self) -> Option<&mut ClipboardData> {
        self.clipboard_data.as_mut()
    }

    /// Set the target's padding-box origin, so `offsetX`/`offsetY` can be
    /// answered. Called by whatever builds the event from a hit test.
    pub fn set_target_origin(&mut self, x: f32, y: f32) {
        self.target_origin = (x, y);
    }

    /// Set the page scroll offset, so `pageX`/`pageY` can be answered.
    pub fn set_scroll_offset(&mut self, x: f32, y: f32) {
        self.scroll_x = x;
        self.scroll_y = y;
    }

    /// Set `screenX`/`screenY`.
    pub fn set_screen_pos(&mut self, x: f32, y: f32) {
        self.screen_x = x;
        self.screen_y = y;
    }
}

impl DomEvent {
    /// `new CustomEvent(type, { detail })`.
    pub fn new_custom(
        event_type: impl Into<String>,
        target: u32,
        detail: impl Into<String>,
    ) -> Self {
        let mut e = Self::new_with_flags(event_type, target, false, false, false);
        e.detail = detail.into();
        e
    }

    /// `CustomEvent.detail`. A string rather than an arbitrary value: this
    /// engine has no script value type to carry.
    pub fn detail(&self) -> &str {
        &self.detail
    }

    /// `CustomEvent.initCustomEvent(...)` — legacy, and like `initEvent` it
    /// does nothing once dispatch has begun.
    pub fn init_custom_event(
        &mut self,
        event_type: &str,
        bubbles: bool,
        cancelable: bool,
        detail: &str,
    ) {
        if self.phase != EventPhase::None {
            return;
        }
        self.init_event(event_type, bubbles, cancelable);
        self.detail = detail.to_string();
    }
}

/// A listener callback.
///
/// It receives the event AND the document, because a listener that cannot
/// touch the DOM is not what `addEventListener` means — a browser handler
/// queries, sets attributes, inserts and removes nodes. The engine hands the
/// document over for the duration of the call by moving the listener map out
/// of it first, so there is no aliasing and no interior-mutability trick at the
/// call site.
pub type EventHandler = Box<dyn FnMut(&mut DomEvent, &mut crate::types::Document) + Send + Sync>;

/// The options half of `addEventListener(type, cb, options)` (DOM §2.7).
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub struct ListenerOptions {
    /// Fire during the capture phase instead of the bubble phase.
    pub capture: bool,
    /// Remove the listener after it fires once.
    pub once: bool,
    /// The listener will not call `preventDefault()`; if it does, the call is
    /// ignored. Lets a scroller run without waiting on the handler.
    pub passive: bool,
}

/// Plain-text ClipboardEvent data. Paste data is read-only; copy/cut handlers can replace it.
#[derive(Clone, Debug, Default)]
pub struct ClipboardData {
    text: Option<String>,
    read_only: bool,
}

impl ClipboardData {
    pub fn get_data(&self, format: &str) -> String {
        if format.eq_ignore_ascii_case("text/plain") || format.eq_ignore_ascii_case("text") {
            self.text.clone().unwrap_or_default()
        } else {
            String::new()
        }
    }
    pub fn set_data(&mut self, format: &str, text: &str) -> bool {
        if self.read_only
            || !(format.eq_ignore_ascii_case("text/plain") || format.eq_ignore_ascii_case("text"))
        {
            return false;
        }
        self.text = Some(text.to_string());
        true
    }
}

impl DomEvent {
    /// An event the system browser delivered, as a listener here receives it.
    pub(crate) fn from_browser(
        event_type: &str,
        target: u32,
        current_target: u32,
        fields: &serde_json::Value,
    ) -> Self {
        let number = |key: &str| fields.get(key).and_then(serde_json::Value::as_f64).unwrap_or(0.0);
        let flag = |key: &str| fields.get(key).and_then(serde_json::Value::as_bool).unwrap_or(false);
        let text = |key: &str| {
            fields.get(key).and_then(serde_json::Value::as_str).unwrap_or("").to_string()
        };
        let mut event = DomEvent::new(event_type, target);
        event.current_target = current_target;
        event.phase = match number("phase") as u16 {
            1 => EventPhase::Capture,
            2 => EventPhase::Target,
            3 => EventPhase::Bubble,
            _ => EventPhase::None,
        };
        event.client_x = number("client_x") as f32;
        event.client_y = number("client_y") as f32;
        event.button = number("button") as u8;
        event.buttons = number("buttons") as u8;
        event.key_code = number("key_code") as u32;
        event.key = text("key");
        event.char_code = {
            let mut chars = event.key.chars();
            match (chars.next(), chars.next()) {
                (Some(c), None) => Some(c),
                _ => None,
            }
        };
        event.code = text("code");
        event.repeat = flag("repeat");
        event.ctrl_key = flag("ctrl_key");
        event.shift_key = flag("shift_key");
        event.alt_key = flag("alt_key");
        event.meta_key = flag("meta_key");
        event.delta_x = number("delta_x") as f32;
        event.delta_y = number("delta_y") as f32;
        event.is_trusted = flag("is_trusted");
        event.time_stamp = number("time_stamp");
        if let Some(path) = fields.get("composed_path").and_then(serde_json::Value::as_array) {
            event.composed_path = path.iter().filter_map(|n| n.as_u64()).map(|n| n as u32).collect();
        }
        event
    }
}

/// `Event.timeStamp`'s origin: a monotonic, comparable stamp.
fn time_origin_elapsed_ms() -> f64 {
    use std::sync::OnceLock;
    use std::time::Instant;
    static ORIGIN: OnceLock<Instant> = OnceLock::new();
    ORIGIN.get_or_init(Instant::now).elapsed().as_secs_f64() * 1000.0
}
