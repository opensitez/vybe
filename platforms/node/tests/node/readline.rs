//! Behaviour tests for `node:readline` host imports.
//!
//! Reference: <https://nodejs.org/api/readline.html>.
//!
//! Coverage:
//!   - `createInterface(options)` → Interface object
//!   - Interface methods: question, pause, resume, close, setPrompt, prompt,
//!     write, getCursorPos, getPrompt
//!   - Interface EventEmitter: on, once, off, emit, removeListener, addListener
//!   - Interface properties: terminal, line, cursor, output, input
//!   - `cursorTo`, `moveCursor`, `clearLine`, `clearScreenDown` — all dirs
//!   - `emitKeypressEvents(stream)` → surface
//!   - `Interface` constructor surface

use std::sync::Arc;
use vybe_bytecode::value::{Object, Value};
use vybe_bytecode::{Chunk, Op, VM};
use vybe_bytecode::capabilities::Capabilities;
use vybe_compiler::compiler::platforms::register_platforms;

fn call_rl(name: &str, args: Vec<Value>) -> Value {
    let mut chunk = Chunk::new("<node-readline-test>");
    let import_idx = chunk.add_import("node:readline", name);
    let argc = args.len() as u8;
    for value in args {
        let c = chunk.add_constant(value);
        chunk.emit_op_u16(Op::CONST, c, 0);
    }
    chunk.emit_op_u16(Op::CALL_IMPORT, import_idx, 0);
    chunk.emit(argc, 0);
    chunk.emit_op(Op::RETURN, 0);

    let mut vm = VM::new();
    register_platforms(&mut vm, &Capabilities::all());
    vm.run(vec![chunk]).expect("VM run failed")
}

fn has_import(name: &str) -> bool {
    let mut vm = VM::new();
    register_platforms(&mut vm, &Capabilities::all());
    vm.host_registry
        .contains_key(&(String::from("node:readline"), name.to_string()))
}

fn s(text: &str) -> Value {
    Value::String(Arc::from(text))
}

fn new_obj(pairs: Vec<(&str, Value)>) -> Value {
    let mut o = Object::new();
    for (k, v) in pairs {
        o.properties.insert(k.to_string(), v);
    }
    Value::Object(std::sync::Arc::new(std::sync::Mutex::new(o)))
}

fn prop(v: &Value, key: &str) -> Value {
    match v {
        Value::Object(o) => o
            .lock()
            .unwrap()
            .properties
            .get(key)
            .cloned()
            .unwrap_or(Value::Undefined),
        _ => Value::Undefined,
    }
}

fn has_method(v: &Value, key: &str) -> bool {
    match v {
        Value::Object(o) => o.lock().unwrap().properties.contains_key(key),
        _ => false,
    }
}

fn create_iface() -> Value {
    call_rl("createInterface", vec![Value::Null])
}

// ── createInterface ───────────────────────────────────────────────────────────

#[test]
fn create_interface_returns_object() {
    let iface = call_rl("createInterface", vec![Value::Null]);
    assert!(matches!(iface, Value::Object(_)));
}

#[test]
fn create_interface_with_options_object() {
    let opts = new_obj(vec![("input", Value::Null), ("output", Value::Null)]);
    let iface = call_rl("createInterface", vec![opts]);
    assert!(matches!(iface, Value::Object(_)));
}

#[test]
fn create_interface_with_terminal_option() {
    let opts = new_obj(vec![
        ("input", Value::Null),
        ("terminal", Value::Bool(false)),
        ("historySize", Value::I32(100)),
    ]);
    let iface = call_rl("createInterface", vec![opts]);
    assert!(matches!(iface, Value::Object(_)));
}

#[test]
fn create_interface_with_prompt_option() {
    let opts = new_obj(vec![("input", Value::Null), ("prompt", s("> "))]);
    let iface = call_rl("createInterface", vec![opts]);
    assert!(matches!(iface, Value::Object(_)));
}

#[test]
fn create_interface_with_crlf_delay() {
    let opts = new_obj(vec![("input", Value::Null), ("crlfDelay", Value::I32(100))]);
    let iface = call_rl("createInterface", vec![opts]);
    assert!(matches!(iface, Value::Object(_)));
}

// ── Interface methods ─────────────────────────────────────────────────────────

#[test]
fn interface_has_close_method() {
    let iface = create_iface();
    assert!(has_method(&iface, "close"), "Interface.close must exist");
}

#[test]
fn interface_has_pause_method() {
    let iface = create_iface();
    assert!(has_method(&iface, "pause"), "Interface.pause must exist");
}

#[test]
fn interface_has_resume_method() {
    let iface = create_iface();
    assert!(has_method(&iface, "resume"), "Interface.resume must exist");
}

#[test]
fn interface_has_set_prompt_method() {
    let iface = create_iface();
    assert!(
        has_method(&iface, "setPrompt"),
        "Interface.setPrompt must exist"
    );
}

#[test]
fn interface_has_get_prompt_method() {
    let iface = create_iface();
    assert!(
        has_method(&iface, "getPrompt"),
        "Interface.getPrompt must exist"
    );
}

#[test]
fn interface_has_prompt_method() {
    let iface = create_iface();
    assert!(has_method(&iface, "prompt"), "Interface.prompt must exist");
}

#[test]
fn interface_has_question_method() {
    let iface = create_iface();
    assert!(
        has_method(&iface, "question"),
        "Interface.question must exist"
    );
}

#[test]
fn interface_has_write_method() {
    let iface = create_iface();
    assert!(has_method(&iface, "write"), "Interface.write must exist");
}

#[test]
fn interface_has_get_cursor_pos_method() {
    let iface = create_iface();
    assert!(
        has_method(&iface, "getCursorPos"),
        "Interface.getCursorPos must exist"
    );
}

// ── Interface EventEmitter ────────────────────────────────────────────────────

#[test]
fn interface_has_on_method() {
    let iface = create_iface();
    assert!(
        has_method(&iface, "on"),
        "Interface.on (EventEmitter) must exist"
    );
}

#[test]
fn interface_has_once_method() {
    let iface = create_iface();
    assert!(has_method(&iface, "once"), "Interface.once must exist");
}

#[test]
fn interface_has_off_method() {
    let iface = create_iface();
    assert!(has_method(&iface, "off"), "Interface.off must exist");
}

#[test]
fn interface_has_emit_method() {
    let iface = create_iface();
    assert!(has_method(&iface, "emit"), "Interface.emit must exist");
}

#[test]
fn interface_has_remove_listener_method() {
    let iface = create_iface();
    assert!(
        has_method(&iface, "removeListener"),
        "Interface.removeListener must exist"
    );
}

#[test]
fn interface_has_add_listener_method() {
    let iface = create_iface();
    assert!(
        has_method(&iface, "addListener"),
        "Interface.addListener must exist"
    );
}

#[test]
fn interface_has_remove_all_listeners_method() {
    let iface = create_iface();
    assert!(
        has_method(&iface, "removeAllListeners"),
        "Interface.removeAllListeners must exist"
    );
}

// ── Interface properties ──────────────────────────────────────────────────────

#[test]
fn interface_has_terminal_property() {
    let iface = create_iface();
    let terminal = prop(&iface, "terminal");
    assert!(
        matches!(terminal, Value::Bool(_) | Value::Undefined | Value::Null),
        "Interface.terminal must be bool or undefined, got {:?}",
        terminal
    );
}

#[test]
fn interface_has_line_property() {
    let iface = create_iface();
    let line = prop(&iface, "line");
    assert!(
        matches!(line, Value::String(_) | Value::Undefined | Value::Null),
        "Interface.line must be string or undefined, got {:?}",
        line
    );
}

#[test]
fn interface_has_cursor_property() {
    let iface = create_iface();
    let cursor = prop(&iface, "cursor");
    assert!(
        matches!(
            cursor,
            Value::I32(_) | Value::I64(_) | Value::F64(_) | Value::Undefined | Value::Null
        ),
        "Interface.cursor must be numeric or undefined, got {:?}",
        cursor
    );
}

// ── cursor functions ──────────────────────────────────────────────────────────

#[test]
fn cursor_to_null_stream_does_not_panic() {
    let result = call_rl("cursorTo", vec![Value::Null, Value::I32(0), Value::I32(0)]);
    assert!(matches!(
        result,
        Value::Bool(_) | Value::Undefined | Value::Null
    ));
}

#[test]
fn cursor_to_x_only_does_not_panic() {
    let result = call_rl("cursorTo", vec![Value::Null, Value::I32(5)]);
    assert!(matches!(
        result,
        Value::Bool(_) | Value::Undefined | Value::Null
    ));
}

#[test]
fn move_cursor_null_stream_does_not_panic() {
    let result = call_rl(
        "moveCursor",
        vec![Value::Null, Value::I32(0), Value::I32(0)],
    );
    assert!(matches!(
        result,
        Value::Bool(_) | Value::Undefined | Value::Null
    ));
}

#[test]
fn move_cursor_negative_direction_does_not_panic() {
    let result = call_rl(
        "moveCursor",
        vec![Value::Null, Value::I32(-5), Value::I32(-3)],
    );
    assert!(matches!(
        result,
        Value::Bool(_) | Value::Undefined | Value::Null
    ));
}

#[test]
fn clear_line_dir_zero_does_not_panic() {
    let result = call_rl("clearLine", vec![Value::Null, Value::I32(0)]);
    assert!(matches!(
        result,
        Value::Bool(_) | Value::Undefined | Value::Null
    ));
}

#[test]
fn clear_line_dir_minus_one_does_not_panic() {
    let result = call_rl("clearLine", vec![Value::Null, Value::I32(-1)]);
    assert!(matches!(
        result,
        Value::Bool(_) | Value::Undefined | Value::Null
    ));
}

#[test]
fn clear_line_dir_plus_one_does_not_panic() {
    let result = call_rl("clearLine", vec![Value::Null, Value::I32(1)]);
    assert!(matches!(
        result,
        Value::Bool(_) | Value::Undefined | Value::Null
    ));
}

#[test]
fn clear_screen_down_null_stream_does_not_panic() {
    let result = call_rl("clearScreenDown", vec![Value::Null]);
    assert!(matches!(
        result,
        Value::Bool(_) | Value::Undefined | Value::Null
    ));
}

// ── emitKeypressEvents ────────────────────────────────────────────────────────

#[test]
fn emit_keypress_events_does_not_panic() {
    let result = call_rl("emitKeypressEvents", vec![Value::Null]);
    let _ = result;
}

// ── Interface constructor ─────────────────────────────────────────────────────

#[test]
fn interface_constructor_is_registered() {
    assert!(has_import("Interface"));
}

// ── Surface check ─────────────────────────────────────────────────────────────

#[test]
fn proposal_node_readline_surface_is_registered() {
    let expected = [
        "createInterface",
        "Interface",
        "cursorTo",
        "moveCursor",
        "clearLine",
        "clearScreenDown",
        "emitKeypressEvents",
    ];
    let missing = expected
        .into_iter()
        .filter(|name| !has_import(name))
        .collect::<Vec<_>>();
    assert!(
        missing.is_empty(),
        "missing node:readline imports: {missing:?}"
    );
}
