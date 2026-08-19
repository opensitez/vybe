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
    WebEngine, WindowOp, WindowValue,
};

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
        repeat: e.repeat,
    }
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
        repeat: f.repeat,
    }
}

struct Widgets;

impl WebEngine for Widgets {
    fn new_document(&self, title: &str) -> DocumentId {
        dom::new_document(title)
    }

    fn new_xml_document(&self, title: &str) -> DocumentId {
        dom::new_xml_document(title)
    }

    fn document(&self, document: DocumentId, op: DomOp) -> DomValue {
        dom::with_document(document, |doc| match op {
            DomOp::CreateElement { tag, input_type } => {
                DomValue::Node(doc.create_element(&tag, &input_type))
            }
            // The bitmap and the box are the same rectangle here — the same
            // simplification `setAttribute("width")` already makes when it
            // writes one by setting the other. A canvas displayed at a
            // different size than its buffer would need them apart.
            DomOp::CanvasSize(node) => match doc.rect(node) {
                Some(rect) => DomValue::Pair(rect.w as f64, rect.h as f64),
                None => DomValue::None,
            },
            DomOp::CreateTextNode(data) => DomValue::Node(doc.create_text_node(&data)),
            DomOp::CreateComment(data) => DomValue::Node(doc.create_comment(&data)),
            DomOp::CreateCDataSection(data) => DomValue::Node(doc.create_cdata_section(&data)),
            DomOp::CreateElementNS {
                namespace,
                qualified_name,
                input_type,
            } => DomValue::Node(doc.create_element_ns(&namespace, &qualified_name, &input_type)),
            DomOp::NamespaceUri(n) => match doc.namespace_uri(n) {
                Some(uri) => DomValue::Text(uri),
                None => DomValue::Null,
            },
            DomOp::Prefix(n) => match doc.prefix(n) {
                Some(prefix) => DomValue::Text(prefix),
                None => DomValue::Null,
            },
            DomOp::LocalName(n) => DomValue::Text(doc.local_name(n)),
            DomOp::CreateProcessingInstruction { target, data } => {
                DomValue::Node(doc.create_processing_instruction(&target, &data))
            }
            DomOp::GetElementById(id) => match doc.get_element_by_id(&id) {
                Some(n) => DomValue::Node(n),
                None => DomValue::Null,
            },
            DomOp::ElementsByTag(tag) => DomValue::Nodes(doc.elements_by_tag(&tag)),
            DomOp::QuerySelector(selectors) => match doc.query_selector(&selectors) {
                Some(node) => DomValue::Node(node),
                None => DomValue::Null,
            },
            DomOp::QuerySelectorAll(selectors) => {
                DomValue::Nodes(doc.query_selector_all(&selectors))
            }
            DomOp::Title => DomValue::Text(doc.title()),
            DomOp::SetTitle(t) => {
                doc.set_title(&t);
                DomValue::None
            }

            DomOp::AppendChild { parent, child } => DomValue::Bool(doc.append_child(parent, child)),
            DomOp::RemoveChild { parent, child } => DomValue::Bool(doc.remove_child(parent, child)),
            DomOp::InsertBefore {
                parent,
                child,
                reference,
            } => DomValue::Bool(doc.insert_before(parent, child, Some(reference))),
            DomOp::ReplaceChild {
                parent,
                new_child,
                old_child,
            } => DomValue::Bool(doc.replace_child(parent, new_child, old_child)),
            DomOp::CloneNode { node, deep } => match doc.clone_node(node, deep) {
                Some(clone) => DomValue::Node(clone),
                None => DomValue::Null,
            },
            DomOp::NodeType(n) => DomValue::Number(f64::from(doc.node_type(n))),
            DomOp::NodeName(n) => DomValue::Text(doc.node_name(n)),
            DomOp::NodeValue(n) => match doc.node_value(n) {
                Some(v) => DomValue::Text(v),
                None => DomValue::Null,
            },
            DomOp::ParentNode(n) => match doc.parent_node(n) {
                Some(p) => DomValue::Node(p),
                None => DomValue::Null,
            },
            DomOp::ChildNodes(n) => DomValue::Nodes(doc.children_of(n)),
            DomOp::InnerHtml(n) => DomValue::Text(doc.inner_html(n)),
            // The SETTER is not here: parsing belongs to the parser, and it
            // needs to re-enter `apply` to build the tree — which would be a
            // second borrow of the document this closure already holds. It is
            // dispatched before the lock instead, in `apply`.
            DomOp::SetInnerHtml { .. } => DomValue::None,
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
                None => DomValue::Null,
            },
            DomOp::AttributeNames(n) => DomValue::Texts(doc.get_attribute_names(n)),
            DomOp::SetAttributeNS {
                node,
                namespace,
                qualified_name,
                value,
            } => {
                doc.set_attribute_ns(node, &namespace, &qualified_name, &value);
                DomValue::None
            }
            DomOp::GetAttributeNS {
                node,
                namespace,
                local_name,
            } => match doc.get_attribute_ns(node, &namespace, &local_name) {
                Some(v) => DomValue::Text(v),
                None => DomValue::Null,
            },
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
            DomOp::SelectedIndex(n) => DomValue::Number(f64::from(doc.selected_index(n))),
            DomOp::SetSelectedIndex(n, i) => {
                doc.set_selected_index(n, i);
                DomValue::None
            }
            DomOp::ItemText(n, i) => DomValue::Text(doc.item_text(n, i)),
            DomOp::SetItemText(n, i, t) => {
                doc.set_item_text(n, i, &t);
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

            // Forwarded whole rather than assembled here: `open` is a
            // reflected attribute, the user-agent stylesheet's rules for a
            // dialog are the toolkit's to apply, and the TOP LAYER a modal
            // enters is a paint-order fact only the renderer can honour.
            // Spelling that out in this file would put a browser's rules
            // outside the engine that draws them.
            DomOp::ShowDialog { node, modal } => {
                doc.show_dialog(node, modal);
                DomValue::None
            }
            DomOp::CloseDialog(node) => {
                doc.close_dialog(node);
                DomValue::None
            }
            DomOp::DialogOpen(node) => DomValue::Bool(doc.dialog_open(node)),

            // `showPicker()` — which picker is decided by the input's TYPE,
            // as HTML decides it. The toolkit's file chooser is a real native
            // dialog, so the pick lands in `value` exactly as a browser's
            // does. A type with no picker returns without doing anything,
            // which is the spec's own answer for a non-picker element — but
            // see the module note: `color` HAS a picker in HTML and does not
            // have one here yet, so that arm is a GAP wearing the same shape.
            DomOp::ShowPicker(node) => {
                let input_type = doc
                    .node(node)
                    .map(|n| n.input_type.clone())
                    .unwrap_or_default();
                if input_type == "file"
                    && let Some(path) = vybe_widgets::dialogs::FileDialog::new("Open").open()
                {
                    doc.set_value(node, &path.to_string_lossy());
                }
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
                None => WindowValue::Null,
            },
            WindowOp::DefaultView(d) => match wnd::default_view(d) {
                Some(w) => WindowValue::Window(w),
                None => WindowValue::Null,
            },
            // The document's own title names the context, which is what
            // `window.name` reads back — `open()` takes the same string as its
            // `target`.
            WindowOp::AdoptTopLevel(d) => WindowValue::Window(wnd::adopt(d, "")),
            WindowOp::Close(w) => {
                wnd::close(w);
                WindowValue::None
            }
            WindowOp::Focus(w) => {
                wnd::focus(w);
                WindowValue::None
            }
            WindowOp::Screen(w) => {
                let (width, height) = wnd::screen(w);
                WindowValue::Pair(width, height)
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
            WindowOp::Name(w) => WindowValue::Text(wnd::name(w)),
            // The message box `vybe_widgets` already has. `alert` is one
            // button, `confirm` is OK/Cancel answering a boolean — which is
            // exactly `DialogChoice::Ok`, so nothing new is drawn or invented.
            WindowOp::Alert(message) => {
                vybe_widgets::dialogs::MessageBox::new(message, "").show();
                WindowValue::None
            }
            WindowOp::Confirm(message) => WindowValue::Bool(matches!(
                vybe_widgets::dialogs::MessageBox::ok_cancel(message, ""),
                vybe_widgets::dialogs::DialogChoice::Ok
            )),
        }
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
                None => EventValue::Null,
            },
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
                    meta_key: st.meta_key,
                }
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
                None => ScheduleValue::Null,
            },
            ScheduleOp::RequestFrame => ScheduleValue::Id(scheduling::frames().request()),
            ScheduleOp::CancelFrame(id) => ScheduleValue::Bool(scheduling::frames().cancel(id)),
            ScheduleOp::TakeDueFrame => match scheduling::frames().take_due() {
                Some(id) => ScheduleValue::Id(id),
                None => ScheduleValue::Null,
            },
            ScheduleOp::TimerDelayMs => match scheduling::timers().delay_until_next_ms() {
                Some(ms) => ScheduleValue::Ms(ms),
                None => ScheduleValue::Null,
            },
            ScheduleOp::FrameDelayMs => match scheduling::frames().delay_until_next_ms() {
                Some(ms) => ScheduleValue::Ms(ms),
                None => ScheduleValue::Null,
            },
            ScheduleOp::Now => ScheduleValue::Ms(scheduling::now_ms()),
        }
    }
}

/// Install the widget toolkit as the web engine.
pub fn install() {
    crate::engine::set_engine(Arc::new(Widgets));
}
