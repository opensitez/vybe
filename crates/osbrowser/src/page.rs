//! The live page: one system-browser tab and the ordered channel to it.
//!
//! Every DOM call becomes one command on the session's queue. Writes are
//! queued and not waited for; a read waits for its answer, and the queue is
//! ordered, so a read always observes every write issued before it. New nodes
//! get their ids here, so creating one costs no round trip either.

use std::collections::{HashMap, VecDeque};
use std::sync::atomic::{AtomicBool, AtomicU8, AtomicU32, AtomicU64, Ordering};
use std::sync::{Arc, Mutex, OnceLock};

use serde::de::DeserializeOwned;
use serde_json::Value;

use crate::{BrowserEvent, BrowserSession};

/// Node ids this process assigns. Above the browser's own sequential range,
/// inside `u32`, so the two never collide.
static NEXT_NODE: AtomicU32 = AtomicU32::new(1 << 30);
static NEXT_PAGE: AtomicU64 = AtomicU64::new(0);
/// The page shown in the tab: the first one opened since the last reset. The
/// pages after it are documents the program holds but does not display.
static VISIBLE_PAGE: AtomicU64 = AtomicU64::new(0);
/// The first document's batch: `IDLE` until one opens, `OPEN` while its
/// construction is being queued, `DONE` once the window loop released it.
static STARTUP_BATCH: AtomicU8 = AtomicU8::new(IDLE);
const IDLE: u8 = 0;
const OPEN: u8 = 1;
const DONE: u8 = 2;
static REPORTED: AtomicBool = AtomicBool::new(false);
/// The process presents no window: pages are hosted by a headless browser.
static HEADLESS: AtomicBool = AtomicBool::new(false);
static SESSION: OnceLock<Result<Arc<BrowserSession>, String>> = OnceLock::new();

pub(crate) fn session() -> Result<&'static Arc<BrowserSession>, String> {
    SESSION
        .get_or_init(|| BrowserSession::start().map(Arc::new))
        .as_ref()
        .map_err(Clone::clone)
}

/// Say once that the browser stopped answering, not once per call: a closed
/// tab fails every DOM call that follows it.
fn report(operation: &str, error: &str) {
    if !REPORTED.swap(true, Ordering::AcqRel) {
        eprintln!("osbrowser: {operation}: {error}");
    }
}

/// Open the tab on first use and mint a page id for a new document.
///
/// The first document's construction is one batch: a program builds its whole
/// UI before the window runs, and sending that as it is built would paint it
/// one element at a time. [`end_startup_batch`] releases it.
pub(crate) fn open_page(title: &str) -> u64 {
    let page = NEXT_PAGE.fetch_add(1, Ordering::Relaxed) + 1;
    let _ = VISIBLE_PAGE.compare_exchange(0, page, Ordering::AcqRel, Ordering::Acquire);
    let Ok(session) = session() else { return page };
    if !session.is_connected() {
        let opened = if HEADLESS.load(Ordering::Acquire) { session.open_headless() } else { session.open() };
        if let Err(error) = opened {
            report("open", &error);
            return page;
        }
    }
    if let Err(error) = session.call(page, "NewDocument", serde_json::json!({ "title": title })) {
        report("NewDocument", &error);
    }
    if STARTUP_BATCH.compare_exchange(IDLE, OPEN, Ordering::AcqRel, Ordering::Acquire).is_ok() {
        session.begin_batch();
    }
    page
}

pub(crate) fn set_headless(headless: bool) {
    HEADLESS.store(headless, Ordering::Release);
}

/// The page the tab displays.
pub(crate) fn visible_page() -> u64 {
    VISIBLE_PAGE.load(Ordering::Acquire)
}

/// Release the startup batch, once. The window loop calls this when it starts.
pub(crate) fn end_startup_batch() {
    if STARTUP_BATCH.compare_exchange(OPEN, DONE, Ordering::AcqRel, Ordering::Acquire).is_ok() {
        if let Ok(session) = session() {
            session.end_batch();
        }
    }
}

/// Ask the page and wait for the answer.
pub(crate) fn read<T: DeserializeOwned + Default>(page: u64, operation: &str, args: Value) -> T {
    match session().and_then(|session| session.call(page, operation, args)) {
        Ok(Value::Null) => T::default(),
        Ok(value) => serde_json::from_value(value).unwrap_or_else(|error| {
            report(operation, &error.to_string());
            T::default()
        }),
        Err(error) => {
            report(operation, &error);
            T::default()
        }
    }
}

/// Queue a mutation. Ordered with every other command; not waited for.
pub(crate) fn write(page: u64, operation: &str, args: Value) {
    if let Err(error) = session().and_then(|session| session.enqueue(page, operation, args, None)) {
        report(operation, &error);
    }
}

/// Queue a node-creating command under an id assigned here.
pub(crate) fn create(page: u64, operation: &str, args: Value) -> u32 {
    let id = NEXT_NODE.fetch_add(1, Ordering::Relaxed);
    match session().and_then(|session| session.enqueue(page, operation, args, Some(u64::from(id)))) {
        Ok(()) => id,
        Err(error) => {
            report(operation, &error);
            0
        }
    }
}

/// Upload raw RGBA bytes in the same ordered stream as DOM commands.
pub(crate) fn write_pixels(page: u64, args: Value, pixels: Vec<u8>) {
    if let Err(error) =
        session().and_then(|session| session.enqueue_payload(page, "canvas_image", args, None, Some(pixels.into())))
    {
        report("canvas_image", &error);
    }
}

pub(crate) fn connected() -> bool {
    SESSION.get().is_some_and(|session| session.as_ref().is_ok_and(|s| s.is_connected()))
}

#[derive(Default)]
struct Routed {
    pages: HashMap<u64, VecDeque<BrowserEvent>>,
}

fn routed() -> &'static Mutex<Routed> {
    static ROUTED: OnceLock<Mutex<Routed>> = OnceLock::new();
    ROUTED.get_or_init(|| Mutex::new(Routed::default()))
}

/// Move what the browser reported into each page's queue, and raw input into
/// `ui_events`, where webcore's window puts it.
pub(crate) fn route() {
    let Some(Ok(session)) = SESSION.get() else { return };
    let events = session.drain_events();
    if events.is_empty() {
        return;
    }
    let Ok(mut routed) = routed().lock() else { return };
    for event in events {
        if event.listener.is_none()
            && matches!(event.kind.as_str(),
                "keydown" | "keyup" | "mousedown" | "mouseup" | "mousemove" | "wheel")
        {
            crate::ui_events::push(ui_event(&event));
        }
        routed.pages.entry(event.document).or_default().push_back(event);
    }
}

/// The events reported for one page since it last asked.
pub(crate) fn take_events(page: u64) -> Vec<BrowserEvent> {
    route();
    routed()
        .lock()
        .ok()
        .and_then(|mut routed| routed.pages.remove(&page))
        .map(Vec::from)
        .unwrap_or_default()
}

fn ui_event(event: &BrowserEvent) -> crate::ui_events::UiEvent {
    let f = &event.fields;
    let int = |key: &str| f.get(key).and_then(Value::as_f64).unwrap_or(0.0) as i32;
    let flag = |key: &str| f.get(key).and_then(Value::as_bool).unwrap_or(false);
    let text = |key: &str| f.get(key).and_then(Value::as_str).unwrap_or("").to_string();
    crate::ui_events::UiEvent {
        kind: event.kind.clone(),
        key: text("key"),
        code: text("code"),
        key_code: int("key_code"),
        client_x: int("client_x"),
        client_y: int("client_y"),
        button: int("button"),
        buttons: int("buttons"),
        delta_y: f.get("delta_y").and_then(Value::as_f64).unwrap_or(0.0),
        ctrl_key: flag("ctrl_key"),
        shift_key: flag("shift_key"),
        alt_key: flag("alt_key"),
        meta_key: flag("meta_key"),
        repeat: flag("repeat"),
    }
}

/// Drop per-run state: queued events, the batch, and the page's documents.
pub(crate) fn reset() {
    if let Ok(mut routed) = routed().lock() {
        *routed = Routed::default();
    }
    end_startup_batch();
    STARTUP_BATCH.store(IDLE, Ordering::Release);
    VISIBLE_PAGE.store(0, Ordering::Release);
    REPORTED.store(false, Ordering::Release);
    if let Some(Ok(session)) = SESSION.get() {
        if session.is_connected() {
            let _ = session.call(0, "Reset", Value::Null);
        }
    }
}
