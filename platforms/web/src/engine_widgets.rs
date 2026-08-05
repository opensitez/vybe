//! The widget toolkit as the engine behind `web:*`.
//!
//! Pure forwarding: every operation is one call into `vybe_widgets`, which is
//! where windows, documents and the event queue actually live. Replacing this
//! file with one that talks to a browser is what swapping engines means —
//! nothing above it changes.

use std::sync::Arc;

use vybe_widgets::dom;
use vybe_widgets::scheduling;
use vybe_widgets::ui_events::{self, UiEvent};
use vybe_widgets::window as wnd;

use crate::engine::{
    DocumentId, DomOp, DomValue, EventOp, EventValue, ScheduleOp, ScheduleValue, UiEventFields,
    WebEngine, WindowOp, WindowValue };

/// Borrow a document — how the window runner reaches the form it renders.
pub fn with_document<T>(id: DocumentId, f: impl FnOnce(&mut dom::Document) -> T) -> Option<T> {
    dom::with_document(id, f)
}

fn to_fields(e: UiEvent) -> UiEventFields {
    UiEventFields {
        kind: e.kind,
        key: e.key,
        code: e.code,
        key_code: e.key_code,
        client_x: e.client_x,
        client_y: e.client_y,
        button: e.button,
        buttons: e.buttons,
        delta_y: e.delta_y,
        ctrl_key: e.ctrl_key,
        shift_key: e.shift_key,
        alt_key: e.alt_key,
        meta_key: e.meta_key,
        repeat: e.repeat }
}

fn from_fields(f: UiEventFields) -> UiEvent {
    UiEvent {
        kind: f.kind,
        key: f.key,
        code: f.code,
        key_code: f.key_code,
        client_x: f.client_x,
        client_y: f.client_y,
        button: f.button,
        buttons: f.buttons,
        delta_y: f.delta_y,
        ctrl_key: f.ctrl_key,
        shift_key: f.shift_key,
        alt_key: f.alt_key,
        meta_key: f.meta_key,
        repeat: f.repeat }
}

struct Widgets;

impl WebEngine for Widgets {
    fn new_document(&self, title: &str) -> DocumentId {
        dom::new_document(title)
    }

    fn document(&self, document: DocumentId, op: DomOp) -> DomValue {
        dom::with_document(document, |doc| match op {
            DomOp::CreateElement { tag, input_type } => {
                DomValue::Node(doc.create_element(&tag, &input_type))
            }
            DomOp::GetElementById(id) => match doc.get_element_by_id(&id) {
                Some(n) => DomValue::Node(n),
                None => DomValue::Null },
            DomOp::ElementsByTag(tag) => DomValue::Nodes(doc.elements_by_tag(&tag)),
            DomOp::Title => DomValue::Text(doc.title()),
            DomOp::SetTitle(t) => {
                doc.set_title(&t);
                DomValue::None
            }

            DomOp::AppendChild { parent, child } => DomValue::Bool(doc.append_child(parent, child)),
            DomOp::RemoveChild { parent, child } => DomValue::Bool(doc.remove_child(parent, child)),
            DomOp::IsConnected(n) => DomValue::Bool(doc.connected(n)),
            DomOp::TextContent(n) => DomValue::Text(doc.text_content(n)),
            DomOp::SetTextContent(n, t) => {
                doc.set_text_content(n, &t);
                DomValue::None
            }

            DomOp::SetAttribute(n, name, value) => {
                doc.set_attribute(n, &name, &value);
                DomValue::None
            }
            DomOp::GetAttribute(n, name) => match doc.get_attribute(n, &name) {
                Some(v) => DomValue::Text(v),
                None => DomValue::Null },
            DomOp::RemoveAttribute(n, name) => {
                doc.remove_attribute(n, &name);
                DomValue::None
            }
            DomOp::SetStyleProperty(n, p, v) => {
                doc.set_style_property(n, &p, &v);
                DomValue::None
            }
            DomOp::GetStyleProperty(n, p) => DomValue::Text(doc.style_property(n, &p)),
            DomOp::Focus(n) => {
                doc.focus(n);
                DomValue::None
            }

            DomOp::Value(n) => DomValue::Text(doc.value(n)),
            DomOp::SetValue(n, v) => {
                doc.set_value(n, &v);
                DomValue::None
            }
            DomOp::Checked(n) => DomValue::Bool(doc.checked(n)),
            DomOp::SetChecked(n, c) => {
                doc.set_checked(n, c);
                DomValue::None
            }
            DomOp::AddItem(n, t) => {
                doc.add_item(n, &t);
                DomValue::None
            }
            DomOp::RemoveItem(n, i) => {
                doc.remove_item(n, i);
                DomValue::None
            }
            DomOp::ClearItems(n) => {
                doc.clear_items(n);
                DomValue::None
            }

            DomOp::DrainEvents => DomValue::Events(
                doc.drain_events()
                    .into_iter()
                    .map(|e| (e.node, e.kind.to_string()))
                    .collect(),
            ),
        })
        .unwrap_or(DomValue::None)
    }

    fn window(&self, op: WindowOp) -> WindowValue {
        match op {
            WindowOp::Open { target, features } => {
                WindowValue::Window(wnd::open(&target, &features))
            }
            WindowOp::Document(w) => match wnd::document(w) {
                Some(d) => WindowValue::Document(d),
                None => WindowValue::Null },
            WindowOp::Close(w) => {
                wnd::close(w);
                WindowValue::None
            }
            WindowOp::Closed(w) => WindowValue::Bool(wnd::closed(w)),
            WindowOp::InnerSize(w) => {
                let (width, height) = wnd::inner_size(w);
                WindowValue::Pair(width, height)
            }
            WindowOp::ResizeTo(w, width, height) => {
                wnd::resize_to(w, width, height);
                WindowValue::None
            }
            WindowOp::MoveTo(w, x, y) => {
                wnd::move_to(w, x, y);
                WindowValue::None
            }
            WindowOp::ScreenPosition(w) => {
                let (x, y) = wnd::screen_position(w);
                WindowValue::Pair(x, y)
            }
            WindowOp::Name(w) => WindowValue::Text(wnd::name(w)) }
    }

    fn events(&self, op: EventOp) -> EventValue {
        let q = ui_events::queue();
        match op {
            EventOp::Dispatch(fields) => {
                q.push(from_fields(fields));
                EventValue::None
            }
            EventOp::Poll => match q.poll() {
                Some(e) => EventValue::Event(to_fields(e)),
                None => EventValue::Null },
            EventOp::Pending => EventValue::Count(q.pending()),
            EventOp::PointerState => {
                let st = q.pointer_state();
                EventValue::Pointer {
                    client_x: st.client_x,
                    client_y: st.client_y,
                    buttons: st.buttons,
                    ctrl_key: st.ctrl_key,
                    shift_key: st.shift_key,
                    alt_key: st.alt_key,
                    meta_key: st.meta_key }
            }
        }
    }

    fn schedule(&self, op: ScheduleOp) -> ScheduleValue {
        match op {
            ScheduleOp::SetTimer(delay_ms) => {
                ScheduleValue::Id(scheduling::timers().schedule(delay_ms))
            }
            ScheduleOp::ClearTimer(id) => ScheduleValue::Bool(scheduling::timers().cancel(id)),
            ScheduleOp::TakeDueTimer => match scheduling::timers().take_due() {
                Some(id) => ScheduleValue::Id(id),
                None => ScheduleValue::Null },
            ScheduleOp::RequestFrame => ScheduleValue::Id(scheduling::frames().request()),
            ScheduleOp::CancelFrame(id) => ScheduleValue::Bool(scheduling::frames().cancel(id)),
            ScheduleOp::TakeDueFrame => match scheduling::frames().take_due() {
                Some(id) => ScheduleValue::Id(id),
                None => ScheduleValue::Null },
            ScheduleOp::TimerDelayMs => match scheduling::timers().delay_until_next_ms() {
                Some(ms) => ScheduleValue::Ms(ms),
                None => ScheduleValue::Null },
            ScheduleOp::FrameDelayMs => match scheduling::frames().delay_until_next_ms() {
                Some(ms) => ScheduleValue::Ms(ms),
                None => ScheduleValue::Null },
            ScheduleOp::Now => ScheduleValue::Ms(scheduling::now_ms()) }
    }
}

/// Install the widget toolkit as the web engine.
pub fn install() {
    crate::engine::set_engine(Arc::new(Widgets));
}
