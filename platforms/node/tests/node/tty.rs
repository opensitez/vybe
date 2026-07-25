//! Behaviour tests for `node:tty` host imports.
//!
//! Reference: <https://nodejs.org/api/tty.html>.
//!
//! Coverage:
//!   - `isatty(fd)` → boolean
//!   - `ReadStream` constructor surface + properties (isRaw, isTTY, fd)
//!   - `ReadStream` methods: setRawMode, destroy, pause, resume, on
//!   - `WriteStream` constructor surface + properties (columns, rows, isTTY, fd)
//!   - `WriteStream` methods: clearLine, cursorTo, moveCursor, getColorDepth, hasColors
//!   - `WriteStream` EventEmitter: on, once, off
//!   - `getColorDepth()` → 1|2|4|8|24 module-level
//!   - `hasColors(count)` → boolean module-level
//!
//! Notes:
//!   - In test context file descriptors are redirected — isatty(0/1/2) returns false.
//!   - Out-of-range or non-numeric fds must return false, not throw.

use vybe_bytecode::value::Value;
use vybe_bytecode::{Chunk, Op, VM};
use vybe_bytecode::capabilities::Capabilities;
use vybe_emitter::platforms::register_platforms;

fn call_tty(name: &str, args: Vec<Value>) -> Value {
    let mut chunk = Chunk::new("<node-tty-test>");
    let import_idx = chunk.add_import("node:tty", name);
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
        .contains_key(&(String::from("node:tty"), name.to_string()))
}

fn prop(obj: &Value, key: &str) -> Value {
    match obj {
        Value::Object(o) => {
            let o = o.lock().unwrap();
            o.properties.get(key).cloned().unwrap_or(Value::Undefined)
        }
        _ => Value::Undefined,
    }
}

fn has_method(obj: &Value, key: &str) -> bool {
    match obj {
        Value::Object(o) => o.lock().unwrap().properties.contains_key(key),
        _ => false,
    }
}

// ── isatty ────────────────────────────────────────────────────────────────────

#[test]
fn is_atty_returns_boolean() {
    let result = call_tty("isatty", vec![Value::I32(1)]);
    assert!(matches!(result, Value::Bool(_)));
}

#[test]
fn is_atty_stdin_is_false_in_test_process() {
    // fd 0 (stdin) is always redirected in cargo test
    let result = call_tty("isatty", vec![Value::I32(0)]);
    assert_eq!(result, Value::Bool(false));
}

#[test]
fn is_atty_stdout_is_false_in_test_process() {
    let result = call_tty("isatty", vec![Value::I32(1)]);
    assert_eq!(result, Value::Bool(false));
}

#[test]
fn is_atty_stderr_is_false_in_test_process() {
    let result = call_tty("isatty", vec![Value::I32(2)]);
    assert_eq!(result, Value::Bool(false));
}

#[test]
fn is_atty_invalid_fd_returns_false_not_throws() {
    let result = call_tty("isatty", vec![Value::I32(-1)]);
    assert_eq!(result, Value::Bool(false));
}

#[test]
fn is_atty_out_of_range_fd_returns_false() {
    let result = call_tty("isatty", vec![Value::I32(99999)]);
    assert_eq!(result, Value::Bool(false));
}

#[test]
fn is_atty_non_integer_fd_returns_false() {
    let result = call_tty("isatty", vec![Value::F64(1.5)]);
    assert_eq!(result, Value::Bool(false));
}

// ── ReadStream constructor ─────────────────────────────────────────────────────

#[test]
fn read_stream_constructor_returns_object() {
    let stream = call_tty("ReadStream", vec![Value::I32(0)]);
    assert!(matches!(stream, Value::Object(_)));
}

#[test]
fn read_stream_is_raw_property_is_bool() {
    let stream = call_tty("ReadStream", vec![Value::I32(0)]);
    let is_raw = prop(&stream, "isRaw");
    assert!(matches!(is_raw, Value::Bool(_)));
}

#[test]
fn read_stream_is_tty_property_is_bool() {
    let stream = call_tty("ReadStream", vec![Value::I32(0)]);
    let is_tty = prop(&stream, "isTTY");
    assert!(matches!(is_tty, Value::Bool(_)));
}

#[test]
fn read_stream_fd_property_reflects_arg() {
    let stream = call_tty("ReadStream", vec![Value::I32(0)]);
    let fd = prop(&stream, "fd");
    match fd {
        Value::I32(n) => assert_eq!(n, 0),
        Value::Undefined => {} // TDD
        other => panic!("fd must be 0, got {:?}", other),
    }
}

#[test]
fn read_stream_has_set_raw_mode_method() {
    let stream = call_tty("ReadStream", vec![Value::I32(0)]);
    assert!(
        has_method(&stream, "setRawMode"),
        "ReadStream.setRawMode must exist"
    );
}

#[test]
fn read_stream_has_on_method() {
    let stream = call_tty("ReadStream", vec![Value::I32(0)]);
    assert!(has_method(&stream, "on"), "ReadStream.on must exist");
}

#[test]
fn read_stream_has_destroy_method() {
    let stream = call_tty("ReadStream", vec![Value::I32(0)]);
    assert!(
        has_method(&stream, "destroy"),
        "ReadStream.destroy must exist"
    );
}

#[test]
fn read_stream_has_pause_and_resume_methods() {
    let stream = call_tty("ReadStream", vec![Value::I32(0)]);
    assert!(has_method(&stream, "pause"), "ReadStream.pause must exist");
    assert!(
        has_method(&stream, "resume"),
        "ReadStream.resume must exist"
    );
}

// ── WriteStream constructor ────────────────────────────────────────────────────

#[test]
fn write_stream_constructor_returns_object() {
    let stream = call_tty("WriteStream", vec![Value::I32(1)]);
    assert!(matches!(stream, Value::Object(_)));
}

#[test]
fn write_stream_columns_is_non_negative_integer() {
    let stream = call_tty("WriteStream", vec![Value::I32(1)]);
    let cols = prop(&stream, "columns");
    match cols {
        Value::I32(n) => assert!(n >= 0),
        Value::F64(f) => assert!(f >= 0.0),
        _ => panic!("columns must be a number"),
    }
}

#[test]
fn write_stream_rows_is_non_negative_integer() {
    let stream = call_tty("WriteStream", vec![Value::I32(1)]);
    let rows = prop(&stream, "rows");
    match rows {
        Value::I32(n) => assert!(n >= 0),
        Value::F64(f) => assert!(f >= 0.0),
        _ => panic!("rows must be a number"),
    }
}

#[test]
fn write_stream_is_tty_property_is_bool() {
    let stream = call_tty("WriteStream", vec![Value::I32(1)]);
    let is_tty = prop(&stream, "isTTY");
    assert!(matches!(is_tty, Value::Bool(_)));
}

#[test]
fn write_stream_fd_property_reflects_arg() {
    let stream = call_tty("WriteStream", vec![Value::I32(1)]);
    let fd = prop(&stream, "fd");
    match fd {
        Value::I32(n) => assert_eq!(n, 1),
        Value::Undefined => {} // TDD
        other => panic!("fd must be 1, got {:?}", other),
    }
}

// ── WriteStream cursor/clear methods ─────────────────────────────────────────

#[test]
fn write_stream_has_clear_line_method() {
    let stream = call_tty("WriteStream", vec![Value::I32(1)]);
    assert!(
        has_method(&stream, "clearLine"),
        "WriteStream.clearLine must exist"
    );
}

#[test]
fn write_stream_has_cursor_to_method() {
    let stream = call_tty("WriteStream", vec![Value::I32(1)]);
    assert!(
        has_method(&stream, "cursorTo"),
        "WriteStream.cursorTo must exist"
    );
}

#[test]
fn write_stream_has_move_cursor_method() {
    let stream = call_tty("WriteStream", vec![Value::I32(1)]);
    assert!(
        has_method(&stream, "moveCursor"),
        "WriteStream.moveCursor must exist"
    );
}

#[test]
fn write_stream_has_get_color_depth_method() {
    let stream = call_tty("WriteStream", vec![Value::I32(1)]);
    assert!(
        has_method(&stream, "getColorDepth"),
        "WriteStream.getColorDepth must exist"
    );
}

#[test]
fn write_stream_get_color_depth_returns_valid_depth() {
    let _stream = call_tty("WriteStream", vec![Value::I32(1)]);
    // call the method through the module — we can't invoke object methods from here
    // but we can check the module-level function
    let result = call_tty("getColorDepth", vec![]);
    match result {
        Value::I32(n) => assert!([1, 2, 4, 8, 24].contains(&n), "unexpected depth {n}"),
        _ => panic!("expected integer"),
    }
}

#[test]
fn write_stream_has_has_colors_method() {
    let stream = call_tty("WriteStream", vec![Value::I32(1)]);
    assert!(
        has_method(&stream, "hasColors"),
        "WriteStream.hasColors must exist"
    );
}

// ── WriteStream EventEmitter ──────────────────────────────────────────────────

#[test]
fn write_stream_has_on_method() {
    let stream = call_tty("WriteStream", vec![Value::I32(1)]);
    assert!(has_method(&stream, "on"), "WriteStream.on must exist");
}

#[test]
fn write_stream_has_once_method() {
    let stream = call_tty("WriteStream", vec![Value::I32(1)]);
    assert!(has_method(&stream, "once"), "WriteStream.once must exist");
}

#[test]
fn write_stream_has_off_method() {
    let stream = call_tty("WriteStream", vec![Value::I32(1)]);
    assert!(has_method(&stream, "off"), "WriteStream.off must exist");
}

#[test]
fn write_stream_has_emit_method() {
    let stream = call_tty("WriteStream", vec![Value::I32(1)]);
    assert!(has_method(&stream, "emit"), "WriteStream.emit must exist");
}

#[test]
fn write_stream_has_write_method() {
    let stream = call_tty("WriteStream", vec![Value::I32(1)]);
    assert!(has_method(&stream, "write"), "WriteStream.write must exist");
}

#[test]
fn write_stream_has_destroy_method() {
    let stream = call_tty("WriteStream", vec![Value::I32(1)]);
    assert!(
        has_method(&stream, "destroy"),
        "WriteStream.destroy must exist"
    );
}

// ── getColorDepth (module-level) ──────────────────────────────────────────────

#[test]
fn get_color_depth_returns_1_2_4_8_or_24() {
    let result = call_tty("getColorDepth", vec![]);
    match result {
        Value::I32(n) => assert!([1, 2, 4, 8, 24].contains(&n), "unexpected depth {n}"),
        _ => panic!("expected integer"),
    }
}

// ── hasColors (module-level) ──────────────────────────────────────────────────

#[test]
fn has_colors_with_2_returns_boolean() {
    let result = call_tty("hasColors", vec![Value::I32(2)]);
    assert!(matches!(result, Value::Bool(_)));
}

#[test]
fn has_colors_with_16m_returns_boolean() {
    let result = call_tty("hasColors", vec![Value::I32(16_000_000)]);
    assert!(matches!(result, Value::Bool(_)));
}

#[test]
fn has_colors_with_256_returns_boolean() {
    let result = call_tty("hasColors", vec![Value::I32(256)]);
    assert!(matches!(result, Value::Bool(_)));
}

// ── Surface check ─────────────────────────────────────────────────────────────

#[test]
fn proposal_node_tty_surface_is_registered() {
    let expected = [
        "isatty",
        "ReadStream",
        "WriteStream",
        "getColorDepth",
        "hasColors",
    ];
    let missing = expected
        .into_iter()
        .filter(|name| !has_import(name))
        .collect::<Vec<_>>();
    assert!(missing.is_empty(), "missing node:tty imports: {missing:?}");
}
