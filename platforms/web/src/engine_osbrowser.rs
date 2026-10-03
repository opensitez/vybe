//! The system browser implementation of the `web:*` engine contract.
//!
//! Node IDs and DOM state live in the browser. The guest VM and its listener
//! callbacks stay in this process; `osbrowser` transports operations and events.

use std::collections::{HashMap, VecDeque};
use std::sync::atomic::{AtomicBool, AtomicU64, Ordering};
use std::sync::{Arc, Mutex, OnceLock};
use std::time::Duration;

use osbrowser::BrowserSession;
use serde::Serialize;
use serde_json::{Value, json};

use crate::engine::{
    DocumentId, DomEventRecord, DomOp, DomValue, EventOp, EventValue, PickerOp, ScheduleOp,
    ScheduleValue, UiEventFields, WebEngine, WindowOp, WindowValue,
};

static NEXT_DOCUMENT: AtomicU64 = AtomicU64::new(0);
// Keep locally allocated handles outside the browser's sequential ID range,
// but inside JavaScript's exact-integer range.
static NEXT_LOCAL_NODE: AtomicU64 = AtomicU64::new(1_000_000_000_000);
static STARTUP_BATCH_USED: AtomicBool = AtomicBool::new(false);
static STARTUP_BATCH_ACTIVE: AtomicBool = AtomicBool::new(false);
static SESSION: OnceLock<Result<Arc<BrowserSession>, String>> = OnceLock::new();

fn session() -> Result<&'static Arc<BrowserSession>, String> {
    SESSION
        .get_or_init(|| BrowserSession::start().map(Arc::new))
        .as_ref()
        .map_err(Clone::clone)
}

pub(crate) fn send(document: DocumentId, operation: &str, args: Value) -> Result<Value, String> {
    session()?.call(document, operation, args)
}

fn split_operation(op: impl Serialize) -> Result<(String, Value), String> {
    match serde_json::to_value(op).map_err(|e| e.to_string())? {
        Value::String(name) => Ok((name, Value::Null)),
        Value::Object(map) if map.len() == 1 => {
            let (name, args) = map.into_iter().next().expect("one variant");
            Ok((name, args))
        }
        _ => Err("invalid engine operation".into()),
    }
}

fn call<T: serde::de::DeserializeOwned>(
    document: DocumentId,
    prefix: &str,
    op: impl Serialize,
) -> Result<T, String> {
    let (name, args) = split_operation(op)?;
    let answer = session()?.call(document, &format!("{prefix}{name}"), args)?;
    serde_json::from_value(answer).map_err(|e| e.to_string())
}

#[derive(Default)]
struct EventQueues {
    dom: HashMap<DocumentId, VecDeque<DomEventRecord>>,
    raw: VecDeque<UiEventFields>,
}

fn event_queues() -> &'static Mutex<EventQueues> {
    static QUEUES: OnceLock<Mutex<EventQueues>> = OnceLock::new();
    QUEUES.get_or_init(|| Mutex::new(EventQueues::default()))
}

fn receive_events() {
    let Ok(session) = session() else { return };
    let Ok(mut queues) = event_queues().lock() else { return };
    for event in session.drain_events() {
        let mut fields = serde_json::from_value::<UiEventFields>(event.fields).unwrap_or_default();
        fields.kind = event.kind.clone();
        let path = if event.path.is_empty() {
            vec![event.node]
        } else {
            event.path.clone()
        };
        for current_target in path {
            queues
                .dom
                .entry(event.document)
                .or_default()
                .push_back(DomEventRecord {
                    target: event.node,
                    current_target,
                    kind: event.kind.clone(),
                    fields: fields.clone(),
                });
        }
        if matches!(event.kind.as_str(),
            "keydown" | "keyup" | "mousedown" | "mouseup" | "mousemove" | "wheel"
        ) {
            queues.raw.push_back(fields);
            if queues.raw.len() > 1024 {
                queues.raw.pop_front();
            }
        }
    }
}

struct OsBrowser;

impl WebEngine for OsBrowser {
    fn new_document(&self, title: &str) -> DocumentId {
        let id = NEXT_DOCUMENT.fetch_add(1, Ordering::Relaxed) + 1;
        let Ok(session) = session() else { return 0 };
        if !session.is_connected() {
            if let Err(error) = session.open() {
                eprintln!("osbrowser: cannot open system browser: {error}");
                return 0;
            }
        }
        if let Err(error) = session.call(id, "NewDocument", json!({ "title": title })) {
            eprintln!("osbrowser: cannot create document: {error}");
            return 0;
        }
        if !STARTUP_BATCH_USED.swap(true, Ordering::AcqRel) {
            session.begin_batch();
            STARTUP_BATCH_ACTIVE.store(true, Ordering::Release);
        }
        id
    }

    fn new_xml_document(&self, title: &str) -> DocumentId {
        let id = self.new_document(title);
        if id != 0 {
            if let Ok(session) = session() {
                let _ = session.call(id, "SetDocumentKind", json!("xml"));
            }
        }
        id
    }

    fn document(&self, document: DocumentId, op: DomOp) -> DomValue {
        if matches!(&op, DomOp::ObserveEvent { .. } | DomOp::UnobserveEvent { .. }) {
            return DomValue::None;
        }
        if matches!(op, DomOp::DrainEvents) {
            receive_events();
            let events = event_queues()
                .lock()
                .ok()
                .and_then(|mut q| q.dom.remove(&document))
                .map(|queue| queue.into_iter().collect())
                .unwrap_or_default();
            return DomValue::Events(events);
        }
        let assigned_node = matches!(&op, DomOp::CreateElement { .. } | DomOp::CreateTextNode(_));
        let queued_mutation = matches!(
            &op,
            DomOp::SetTitle(_)
                | DomOp::AppendChild { .. }
                | DomOp::RemoveChild { .. }
                | DomOp::ReplaceChild { .. }
                | DomOp::SetInnerHtml { .. }
                | DomOp::SetOuterHtml { .. }
                | DomOp::InsertAdjacentHtml { .. }
                | DomOp::SetTextContent(..)
                | DomOp::SetAttribute(..)
                | DomOp::RemoveAttribute(..)
                | DomOp::SetAttributeNS { .. }
                | DomOp::SetStyleProperty(..)
                | DomOp::Focus(_)
                | DomOp::SetValue(..)
                | DomOp::SetChecked(..)
                | DomOp::SetSelectedIndex(..)
                | DomOp::SetItemText(..)
                | DomOp::AddItem(..)
                | DomOp::RemoveItem(..)
                | DomOp::ClearItems(_)
                | DomOp::ShowDialog { .. }
                | DomOp::CloseDialog(_)
        );
        if assigned_node || queued_mutation {
            let queued = split_operation(&op).and_then(|(name, args)| {
                let id = assigned_node.then(|| NEXT_LOCAL_NODE.fetch_add(1, Ordering::Relaxed));
                session()?.enqueue(document, &name, args, id)?;
                Ok(id)
            });
            return match queued {
                Ok(Some(id)) => DomValue::Node(id),
                Ok(None) => DomValue::None,
                Err(error) => {
                    eprintln!("osbrowser DOM: {error}");
                    DomValue::Null
                }
            };
        }
        match call(document, "", op) {
            Ok(value) => value,
            Err(error) => {
                eprintln!("osbrowser DOM: {error}");
                DomValue::Null
            }
        }
    }

    fn window(&self, op: WindowOp) -> WindowValue {
        match op {
            WindowOp::Open { target, .. } => WindowValue::Window(self.new_document(&target)),
            WindowOp::Document(window) => WindowValue::Document(window),
            WindowOp::DefaultView(document) | WindowOp::AdoptTopLevel(document) => {
                WindowValue::Window(document)
            }
            other => call(0, "Window", other).unwrap_or(WindowValue::Null),
        }
    }

    fn events(&self, op: EventOp) -> EventValue {
        receive_events();
        match op {
            EventOp::Poll => event_queues()
                .lock()
                .ok()
                .and_then(|mut q| q.raw.pop_front())
                .map(EventValue::Event)
                .unwrap_or(EventValue::Null),
            EventOp::Pending => EventValue::Count(
                event_queues().lock().map(|q| q.raw.len()).unwrap_or(0),
            ),
            EventOp::Dispatch(event) => {
                if let Ok(mut q) = event_queues().lock() {
                    q.raw.push_back(event);
                }
                EventValue::None
            }
            EventOp::PointerState => call(0, "Event", EventOp::PointerState)
                .unwrap_or(EventValue::Null),
        }
    }

    fn schedule(&self, op: ScheduleOp) -> ScheduleValue {
        call(0, "Schedule", op).unwrap_or(ScheduleValue::Null)
    }

    fn picker(&self, op: PickerOp) -> Vec<String> {
        let paths = match op {
            PickerOp::Open { title, filters, directory, multiple } => {
                let mut dialog = rfd::FileDialog::new().set_title(&title);
                if !directory.is_empty() {
                    dialog = dialog.set_directory(&directory);
                }
                for (name, extensions) in &filters {
                    let extensions: Vec<&str> = extensions.iter().map(String::as_str).collect();
                    dialog = dialog.add_filter(name, &extensions);
                }
                if multiple {
                    dialog.pick_files().unwrap_or_default()
                } else {
                    dialog.pick_file().into_iter().collect()
                }
            }
            PickerOp::Save { title, filters, directory, suggested } => {
                let mut dialog = rfd::FileDialog::new().set_title(&title);
                if !directory.is_empty() {
                    dialog = dialog.set_directory(&directory);
                }
                if !suggested.is_empty() {
                    dialog = dialog.set_file_name(&suggested);
                }
                for (name, extensions) in &filters {
                    let extensions: Vec<&str> = extensions.iter().map(String::as_str).collect();
                    dialog = dialog.add_filter(name, &extensions);
                }
                dialog.save_file().into_iter().collect()
            }
            PickerOp::Directory { title, directory } => {
                let mut dialog = rfd::FileDialog::new().set_title(&title);
                if !directory.is_empty() {
                    dialog = dialog.set_directory(&directory);
                }
                dialog.pick_folder().into_iter().collect()
            }
        };
        paths.into_iter().map(|path| path.to_string_lossy().into_owned()).collect()
    }
}

pub fn install() {
    crate::engine::set_engine(Arc::new(OsBrowser));
}

pub fn reset() {
    finish_startup_batch();
    STARTUP_BATCH_USED.store(false, Ordering::Release);
    if let Ok(mut queues) = event_queues().lock() {
        *queues = EventQueues::default();
    }
    if let Some(Ok(session)) = SESSION.get() {
        if session.is_connected() {
            let _ = session.call(0, "Reset", Value::Null);
        }
    }
}

pub fn connected() -> bool {
    session().is_ok_and(|session| session.is_connected())
}

pub fn wait_for_event(timeout: Duration) {
    if let Ok(session) = session() {
        session.wait_for_event(timeout);
    }
}

pub fn begin_batch() {
    if let Ok(session) = session() {
        session.begin_batch();
    }
}

pub fn end_batch() {
    if let Ok(session) = session() {
        session.end_batch();
    }
}

pub fn finish_startup_batch() {
    if STARTUP_BATCH_ACTIVE.swap(false, Ordering::AcqRel) {
        end_batch();
    }
}
