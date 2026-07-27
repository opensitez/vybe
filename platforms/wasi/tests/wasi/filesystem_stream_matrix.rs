use std::path::PathBuf;
use std::sync::atomic::{AtomicU64, Ordering};

use vybe_bytecode::value::{ObjectKind, Value};
use vybe_bytecode::{Chunk, Op, VM};
use vybe_bytecode::capabilities::Capabilities;
use vybe_compiler::compiler::platforms::register_platforms;

fn scratch_dir(label: &str) -> PathBuf {
    static COUNTER: AtomicU64 = AtomicU64::new(0);
    let id = COUNTER.fetch_add(1, Ordering::SeqCst);
    let dir = std::env::temp_dir().join(format!(
        "vybe-wasi-fs-stream-matrix-test-{}-{}-{}",
        std::process::id(),
        label,
        id
    ));
    let _ = std::fs::remove_dir_all(&dir);
    std::fs::create_dir_all(&dir).expect("scratch dir mkdir");
    dir
}

fn invoke(module: &str, name: &str, args: Vec<Value>) -> Value {
    let mut chunk = Chunk::new("<wasi-fs-stream-matrix-test>");
    let import_idx = chunk.add_import(module, name);
    let argc = args.len() as u8;
    for value in args {
        let constant = chunk.add_constant(value);
        chunk.emit_op_u16(Op::CONST, constant, 0);
    }
    chunk.emit_op_u16(Op::CALL_IMPORT, import_idx, 0);
    chunk.emit(argc, 0);
    chunk.emit_op(Op::RETURN, 0);

    let mut vm = VM::new();
    register_platforms(&mut vm, &Capabilities::all());
    vm.run(vec![chunk]).expect("VM run failed")
}

fn types(name: &str, args: Vec<Value>) -> Value {
    invoke("wasi:filesystem/types", name, args)
}

fn s(text: &str) -> Value {
    Value::String(std::sync::Arc::from(text))
}

fn open_test_root(dir: &PathBuf) -> Value {
    invoke(
        "wasi:filesystem/types",
        "__test_open_root",
        vec![s(dir.to_str().unwrap())],
    )
}

fn bytes_to_vec(value: &Value) -> Vec<u8> {
    let Value::Object(object) = value else {
        return Vec::new();
    };
    let object = object.lock().unwrap();
    let ObjectKind::Array(bytes) = &object.kind else {
        return Vec::new();
    };
    bytes
        .iter()
        .filter_map(|value| match value {
            Value::I32(byte) => Some(*byte as u8),
            Value::F64(byte) => Some(*byte as u8),
            _ => None,
        })
        .collect()
}

fn open_stream(offset: f64) -> Value {
    let dir = scratch_dir("stream_open");
    std::fs::write(dir.join("payload.bin"), b"abcdef").unwrap();
    let root = open_test_root(&dir);
    let descriptor = types(
        "[method]descriptor.open-at",
        vec![
            root,
            Value::I32(0),
            s("payload.bin"),
            Value::I32(0),
            Value::I32(0),
        ],
    );
    types(
        "[method]descriptor.read-via-stream",
        vec![descriptor, Value::F64(offset)],
    )
}

macro_rules! read_case {
    ($name:ident, $method:expr, $offset:expr, $length:expr, $expected:expr) => {
        #[test]
        fn $name() {
            let stream = open_stream($offset);
            let result = invoke(
                "wasi:io/streams",
                $method,
                vec![stream, Value::F64($length)],
            );
            assert_eq!(bytes_to_vec(&result), $expected);
        }
    };
}

read_case!(
    read_from_offset_zero_returns_first_byte,
    "[method]input-stream.read",
    0.0,
    1.0,
    b"a"
);
read_case!(
    read_from_offset_one_returns_middle_slice,
    "[method]input-stream.read",
    1.0,
    2.0,
    b"bc"
);
read_case!(
    read_from_offset_four_returns_remaining_suffix,
    "[method]input-stream.read",
    4.0,
    8.0,
    b"ef"
);

read_case!(
    blocking_read_from_offset_zero_returns_first_byte,
    "[method]input-stream.blocking-read",
    0.0,
    1.0,
    b"a"
);
read_case!(
    blocking_read_from_offset_two_returns_middle_slice,
    "[method]input-stream.blocking-read",
    2.0,
    2.0,
    b"cd"
);
read_case!(
    blocking_read_from_offset_four_returns_remaining_suffix,
    "[method]input-stream.blocking-read",
    4.0,
    8.0,
    b"ef"
);

read_case!(
    read_from_exact_end_returns_empty_array,
    "[method]input-stream.read",
    6.0,
    4.0,
    b""
);
read_case!(
    blocking_read_from_exact_end_returns_empty_array,
    "[method]input-stream.blocking-read",
    6.0,
    4.0,
    b""
);
read_case!(
    read_from_beyond_end_returns_empty_array,
    "[method]input-stream.read",
    12.0,
    4.0,
    b""
);
read_case!(
    blocking_read_from_beyond_end_returns_empty_array,
    "[method]input-stream.blocking-read",
    12.0,
    4.0,
    b""
);

#[test]
fn sequential_reads_advance_stream_position() {
    let stream = open_stream(0.0);
    let first = invoke(
        "wasi:io/streams",
        "[method]input-stream.read",
        vec![stream.clone(), Value::F64(1.0)],
    );
    let second = invoke(
        "wasi:io/streams",
        "[method]input-stream.read",
        vec![stream.clone(), Value::F64(2.0)],
    );
    let third = invoke(
        "wasi:io/streams",
        "[method]input-stream.read",
        vec![stream, Value::F64(8.0)],
    );
    assert_eq!(bytes_to_vec(&first), b"a");
    assert_eq!(bytes_to_vec(&second), b"bc");
    assert_eq!(bytes_to_vec(&third), b"def");
}
