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

use vybe_runtime::vm::{HostFnDecl, ResourceBinding, ResourceMemberKind};
use vybe_runtime::{FuncSig, HostContext, VM, ValType, Value};

use crate::engine::{WindowId, WindowOp, WindowValue, window};

/// A `Window` is a RESOURCE — the engine owns the browsing context, the guest
/// holds a handle. Every function below except `open` is a method on it.
const WINDOW: &str = "window";

/// A borrowed window handle. Borrowed, not owned: reading `innerWidth` does not
/// consume the window, and `close()` does not either — a closed window is still
/// a valid handle that answers `closed = true`, which is the spec's own model.
fn win() -> ValType {
    ValType::Borrow(WINDOW.to_string())
}

/// Register a `web:window` method WITH its signature, in one call.
///
/// Same reason as `platforms/web/src/html.rs`'s `dom_fn`: a signature written
/// beside the registration is a second statement of one fact, and the two
/// drift.
fn win_fn(
    vm: &mut VM,
    name: &str,
    kebab: &str,
    params: Vec<ValType>,
    results: Vec<ValType>,
    call: Box<dyn Fn(&mut HostContext, &[Value]) -> Value + Send + Sync>,
) {
    vm.register_host(
        HostFnDecl::new("web:window", name, call)
            .with_sig(FuncSig {
                name: kebab.to_string(),
                params,
                results,
            })
            .method_on(WINDOW),
    );
}

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
        _ => 0.0,
    }
}

fn second(v: WindowValue) -> f64 {
    match v {
        WindowValue::Pair(_, b) => b,
        _ => 0.0,
    }
}

pub fn register(vm: &mut VM) {
    // window.open(url, target, features) → Window
    //
    // The one STATIC: it is how a browsing context comes into existence, so it
    // has no window to be a method on and returns an OWNED handle. Everything
    // below borrows the handle this produced.
    vm.register_host(
        HostFnDecl::new(
            "web:window",
            "open",
            Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
            // `url` is accepted and ignored: there is no navigation here, so
            // every window opens the spec's initial `about:blank`.
                match window(WindowOp::Open {
                    target: str_arg(args, 1),
                    features: str_arg(args, 2),
                }) {
                    WindowValue::Window(id) => Value::F64(id as f64),
                    _ => Value::Null,
                }
            }),
        )
        .with_sig(FuncSig {
            name: "open".to_string(),
            // `url` is param 0 and deliberately unread — declared because the
            // CALLER passes it and the spec has it, not because this reads it.
            params: vec![ValType::String, ValType::String, ValType::String],
            results: vec![ValType::Own(WINDOW.to_string())],
        })
        .resource_member(ResourceBinding {
            resource: WINDOW.to_string(),
            kind: ResourceMemberKind::Static,
            // A static has no self to borrow — this is how you GET a window.
            borrows_self: false,
        }),
    );

    // window.document → the handle every element call is scoped to
    win_fn(
        vm,
        "document",
        "document",
        vec![win()],
        vec![ValType::Borrow("document".to_string())],
        Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
            match window(WindowOp::Document(win_arg(args, 0))) {
                WindowValue::Document(d) => Value::F64(d as f64),
                _ => Value::Null,
            }
        }),
    );

    // window.close() / window.closed
    win_fn(
        vm,
        "close",
        "close",
        vec![win()],
        vec![],
        Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
            window(WindowOp::Close(win_arg(args, 0)));
            Value::Null
        }),
    );
    win_fn(
        vm,
        "closed",
        "closed",
        vec![win()],
        vec![ValType::Bool],
        Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
            // An unknown handle reads closed, as a stale reference does.
            Value::Bool(!matches!(
                window(WindowOp::Closed(win_arg(args, 0))),
                WindowValue::Bool(false)
            ))
        }),
    );

    // window.innerWidth / innerHeight
    //
    // `f64`, not `i32`: these are the document's viewport read back, and CSS
    // pixels are fractional under a zoom or a device pixel ratio. The IDL says
    // `long`, but rounding here would make the two disagree with the CSSOM
    // geometry next door, which is text with units and never rounded.
    win_fn(
        vm,
        "innerWidth",
        "inner-width",
        vec![win()],
        vec![ValType::F64],
        Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
            Value::F64(first(window(WindowOp::InnerSize(win_arg(args, 0)))))
        }),
    );
    win_fn(
        vm,
        "innerHeight",
        "inner-height",
        vec![win()],
        vec![ValType::F64],
        Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
            Value::F64(second(window(WindowOp::InnerSize(win_arg(args, 0)))))
        }),
    );

    // window.screenX / screenY
    win_fn(
        vm,
        "screenX",
        "screen-x",
        vec![win()],
        vec![ValType::F64],
        Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
            Value::F64(first(window(WindowOp::ScreenPosition(win_arg(args, 0)))))
        }),
    );
    win_fn(
        vm,
        "screenY",
        "screen-y",
        vec![win()],
        vec![ValType::F64],
        Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
            Value::F64(second(window(WindowOp::ScreenPosition(win_arg(args, 0)))))
        }),
    );

    // window.resizeTo(width, height) / moveTo(x, y)
    win_fn(
        vm,
        "resizeTo",
        "resize-to",
        vec![win(), ValType::F64, ValType::F64],
        vec![],
        Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
            window(WindowOp::ResizeTo(
                win_arg(args, 0),
                num_arg(args, 1),
                num_arg(args, 2),
            ));
            Value::Null
        }),
    );
    win_fn(
        vm,
        "moveTo",
        "move-to",
        vec![win(), ValType::F64, ValType::F64],
        vec![],
        Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
            window(WindowOp::MoveTo(
                win_arg(args, 0),
                num_arg(args, 1),
                num_arg(args, 2),
            ));
            Value::Null
        }),
    );

    // window.alert(message) / window.confirm(message)
    //
    // Spelled without a `Window` argument on purpose: both are called on the
    // global object in every real page (`alert("hi")`), and a guest that never
    // opened a window still has one to talk to.
    // They are therefore STATIC, not methods: their first parameter is the
    // MESSAGE, not a window. Declaring them like the rest would have said the
    // string was a handle — the one place in this file where copying the
    // surrounding shape would have been wrong.
    vm.register_host(
        HostFnDecl::new(
            "web:window",
            "alert",
            Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
                window(WindowOp::Alert(str_arg(args, 0)));
                Value::Null
            }),
        )
        .with_sig(FuncSig {
            name: "alert".to_string(),
            params: vec![ValType::String],
            results: vec![],
        })
        .resource_member(ResourceBinding {
            resource: WINDOW.to_string(),
            kind: ResourceMemberKind::Static,
            borrows_self: false,
        }),
    );
    vm.register_host(
        HostFnDecl::new(
            "web:window",
            "confirm",
            Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
                match window(WindowOp::Confirm(str_arg(args, 0))) {
                    WindowValue::Bool(ok) => Value::Bool(ok),
                    _ => Value::Bool(false),
                }
            }),
        )
        .with_sig(FuncSig {
            name: "confirm".to_string(),
            params: vec![ValType::String],
            results: vec![ValType::Bool],
        })
        .resource_member(ResourceBinding {
            resource: WINDOW.to_string(),
            kind: ResourceMemberKind::Static,
            borrows_self: false,
        }),
    );

    // window.name
    win_fn(
        vm,
        "name",
        "name",
        vec![win()],
        vec![ValType::String],
        Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
            match window(WindowOp::Name(win_arg(args, 0))) {
                WindowValue::Text(s) => Value::String(s.into()),
                _ => Value::String("".into()),
            }
        }),
    );
}
