//! WHATWG HTML §7 — the `web:window` host functions.
//!
//! Exposure only. Browsing contexts live in the engine (see
//! [`engine`](crate::engine)), because a window and the document it holds are
//! one thing and splitting them across crates is how their sizes drift apart.
//! `innerWidth`/`innerHeight` are the document's viewport read back, not a
//! second copy.
//!
//! `open(url, target, features)` creates a browsing context AND its initial
//! `about:blank` document, then hands back the `Window`. That is the spec's
//! own bootstrap: a page builds a new window's contents by opening it and
//! calling `createElement` **on that window's document**. Nothing here
//! invents a `createWindow`.
//!
//! One honest difference from a browser: there, `open` is a method on an
//! existing `Window`, because the user agent already made a tab. A wasm guest
//! has no tab — it may legitimately have no window at all — so the first
//! `open` comes from the namespace. Everything after it is standard.

use vybe_runtime::{HostContext, VM, Value};

use crate::engine::{window, WindowId, WindowOp, WindowValue};

fn num_arg(args: &[Value], idx: usize) -> f64 {
    args.get(idx).map(|v| v.as_f64()).unwrap_or(0.0)
}

fn win_arg(args: &[Value], idx: usize) -> WindowId {
    args.get(idx).map(|v| v.as_f64() as WindowId).unwrap_or(0)
}

fn str_arg(args: &[Value], idx: usize) -> String {
    args.get(idx)
        .map(|v| format!("{}", v))
        .filter(|s| s != "null" && s != "undefined")
        .unwrap_or_default()
}

/// The first of a `(width, height)` / `(x, y)` pair; `second` takes the other.
fn first(v: WindowValue) -> f64 {
    match v {
        WindowValue::Pair(a, _) => a,
        _ => 0.0 }
}

fn second(v: WindowValue) -> f64 {
    match v {
        WindowValue::Pair(_, b) => b,
        _ => 0.0 }
}

pub fn register(vm: &mut VM) {
    // window.open(url, target, features) → Window
    vm.register_host_fn(
        "web:window",
        "open",
        Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
            // `url` is accepted and ignored: there is no navigation here, so
            // every window opens the spec's initial `about:blank`.
            match window(WindowOp::Open {
                target: str_arg(args, 1),
                features: str_arg(args, 2) }) {
                WindowValue::Window(id) => Value::F64(id as f64),
                _ => Value::Null }
        }),
    );

    // window.document → the handle every element call is scoped to
    vm.register_host_fn(
        "web:window",
        "document",
        Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
            match window(WindowOp::Document(win_arg(args, 0))) {
                WindowValue::Document(d) => Value::F64(d as f64),
                _ => Value::Null }
        }),
    );

    // window.close() / window.closed
    vm.register_host_fn(
        "web:window",
        "close",
        Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
            window(WindowOp::Close(win_arg(args, 0)));
            Value::Null
        }),
    );
    vm.register_host_fn(
        "web:window",
        "closed",
        Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
            // An unknown handle reads closed, as a stale reference does.
            Value::Bool(!matches!(
                window(WindowOp::Closed(win_arg(args, 0))),
                WindowValue::Bool(false)
            ))
        }),
    );

    // window.innerWidth / innerHeight
    vm.register_host_fn(
        "web:window",
        "innerWidth",
        Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
            Value::F64(first(window(WindowOp::InnerSize(win_arg(args, 0)))))
        }),
    );
    vm.register_host_fn(
        "web:window",
        "innerHeight",
        Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
            Value::F64(second(window(WindowOp::InnerSize(win_arg(args, 0)))))
        }),
    );

    // window.screenX / screenY
    vm.register_host_fn(
        "web:window",
        "screenX",
        Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
            Value::F64(first(window(WindowOp::ScreenPosition(win_arg(args, 0)))))
        }),
    );
    vm.register_host_fn(
        "web:window",
        "screenY",
        Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
            Value::F64(second(window(WindowOp::ScreenPosition(win_arg(args, 0)))))
        }),
    );

    // window.resizeTo(width, height) / moveTo(x, y)
    vm.register_host_fn(
        "web:window",
        "resizeTo",
        Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
            window(WindowOp::ResizeTo(
                win_arg(args, 0),
                num_arg(args, 1),
                num_arg(args, 2),
            ));
            Value::Null
        }),
    );
    vm.register_host_fn(
        "web:window",
        "moveTo",
        Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
            window(WindowOp::MoveTo(
                win_arg(args, 0),
                num_arg(args, 1),
                num_arg(args, 2),
            ));
            Value::Null
        }),
    );

    // window.name
    vm.register_host_fn(
        "web:window",
        "name",
        Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
            match window(WindowOp::Name(win_arg(args, 0))) {
                WindowValue::Text(s) => Value::String(s.into()),
                _ => Value::String("".into()) }
        }),
    );
}
