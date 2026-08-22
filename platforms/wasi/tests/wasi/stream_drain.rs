//! Draining a WASI 0.3.1 `read-via-stream` the way a guest actually does it.
//!
//! 0.3.1 replaced the 0.2 pair `read-via-stream(offset) -> input-stream` +
//! `wasi:io/streams.[method]input-stream.read` with
//!
//!     read-via-stream: func(offset: filesize)
//!         -> tuple<stream<u8>, future<result<_, error-code>>>
//!
//! where element 0 is a Component Model stream end drained with
//! `canon stream.read`. `wasi:io` does not exist in 0.3.1, so a test that
//! reaches for `input-stream.read` is asserting a deleted package.
//!
//! The drain has to happen inside the SAME `vm.run` as the open: a stream end
//! is an index into that VM's handle table, so handing the value out and
//! calling back in gives a handle from a VM that is gone. That is why this is
//! one chunk rather than the one-host-call-per-VM shape the other helpers use.
//!
//! It calls the compiler's own `emit_read_stream_to_bytes`, not a
//! reimplementation of it — the point of these tests is that the real guest
//! sequence works against the real host function.

use std::path::Path;

use vybe_compiler::primitives::platforms::register_platforms;
use vybe_runtime::capabilities::Capabilities;
use vybe_runtime::value::{ObjectKind, Value};
use vybe_runtime::{Chunk, Op, VM};

/// Open `dir/name` and drain it from `offset` to EOF.
pub fn read_via_stream(dir: &Path, name: &str, offset: f64) -> Vec<u8> {
    let mut chunk = Chunk::new("<wasi-fs-stream-drain>");

    let preopens = chunk.add_import("wasi:filesystem/preopens", "get-directories");
    let at = chunk.add_import("ecma:array", "at");
    let open_at = chunk.add_import("wasi:filesystem/types", "[method]descriptor.open-at");
    let read_via = chunk.add_import(
        "wasi:filesystem/types",
        "[method]descriptor.read-via-stream",
    );

    // The preopen descriptor: get-directories()[0][0].
    chunk.emit_call(preopens, 0, 0);
    chunk.emit_i32_const(0, 0);
    chunk.emit_call(at, 2, 0);
    chunk.emit_i32_const(0, 0);
    chunk.emit_call(at, 2, 0);

    chunk.emit_i32_const(0, 0); // path-flags
    chunk.emit_string_const(&format!("{}/{}", dir.display(), name), 0);
    chunk.emit_i32_const(0, 0); // open-flags: neither create nor truncate
    chunk.emit_i32_const(1, 0); // descriptor-flags: read
    chunk.emit_call(open_at, 5, 0);

    chunk.emit_f64_const(offset, 0);
    chunk.emit_call(read_via, 2, 0);
    // tuple<stream<u8>, future<…>> — element 0 is the readable end.
    chunk.emit_i32_const(0, 0);
    chunk.emit_call(at, 2, 0);
    vybe_compiler::primitives::io::emit_read_stream_to_bytes(&mut chunk, 0);
    chunk.emit_op(Op::RETURN, 0);

    let mut vm = VM::new();
    register_platforms(&mut vm, &Capabilities::all());
    let result = vm.run(vec![chunk]).expect("VM run failed");
    bytes_to_vec(&result)
}

/// The `future<result<_, error-code>>` half of the same tuple, so a test can
/// assert the outcome separately from the bytes.
pub fn read_via_stream_outcome(dir: &Path, name: &str, offset: f64) -> Value {
    let mut chunk = Chunk::new("<wasi-fs-stream-outcome>");

    let preopens = chunk.add_import("wasi:filesystem/preopens", "get-directories");
    let at = chunk.add_import("ecma:array", "at");
    let open_at = chunk.add_import("wasi:filesystem/types", "[method]descriptor.open-at");
    let read_via = chunk.add_import(
        "wasi:filesystem/types",
        "[method]descriptor.read-via-stream",
    );

    chunk.emit_call(preopens, 0, 0);
    chunk.emit_i32_const(0, 0);
    chunk.emit_call(at, 2, 0);
    chunk.emit_i32_const(0, 0);
    chunk.emit_call(at, 2, 0);
    chunk.emit_i32_const(0, 0);
    chunk.emit_string_const(&format!("{}/{}", dir.display(), name), 0);
    chunk.emit_i32_const(0, 0);
    chunk.emit_i32_const(1, 0);
    chunk.emit_call(open_at, 5, 0);
    chunk.emit_f64_const(offset, 0);
    chunk.emit_call(read_via, 2, 0);
    chunk.emit_i32_const(1, 0);
    chunk.emit_call(at, 2, 0);
    chunk.emit_op(Op::RETURN, 0);

    let mut vm = VM::new();
    register_platforms(&mut vm, &Capabilities::all());
    vm.run(vec![chunk]).expect("VM run failed")
}

/// Enumerate `dir` through `read-directory`, as `(type, name)` pairs sorted by
/// name.
///
/// The `stream<directory-entry>` sibling of [`read_via_stream`], and it exists
/// for the same one-chunk-one-VM reason — but it is worth saying what changes
/// when the element stops being a byte. `canon stream.read` copies ELEMENTS at
/// their canonical stride, so lifting one means reading a `descriptor-type`
/// discriminant and a (ptr, length) string pair back out of linear memory at
/// offsets the record layout fixes. That lift is
/// `fs_path::emit_read_directory_entries`, and this calls it rather than
/// repeating it: a reader only the test suite owns would prove that the HOST
/// side works while every language still could not list a directory, which is
/// the `__test_open_root` mistake wearing a different hat.
///
/// Sorted here because `read-directory` promises no order — the WIT says
/// nothing about it and the host hands over whatever `read_dir` yielded, so an
/// assertion on sequence would be asserting the filesystem's mood.
pub fn read_directory(dir: &Path) -> Vec<(String, String)> {
    let mut chunk = Chunk::new("<wasi-fs-read-directory>");
    chunk.emit_string_const(&dir.display().to_string(), 0);
    vybe_compiler::primitives::fs_path::emit_read_directory_entries(&mut chunk, 0);
    chunk.emit_op(Op::RETURN, 0);

    let mut vm = VM::new();
    register_platforms(&mut vm, &Capabilities::all());
    let result = vm.run(vec![chunk]).expect("VM run failed");

    let Value::Object(object) = &result else {
        panic!("expected an array of directory entries, got {result:?}");
    };
    let object = object.lock().unwrap();
    let ObjectKind::Array(entries) = &object.kind else {
        panic!("expected an array, got {:?}", object.kind);
    };
    let mut out: Vec<(String, String)> = entries.iter().map(entry_pair).collect();
    // By NAME — sorting the pair would order by `type` first, which reads as a
    // stable order right up until a test's expectation lists a directory after
    // a file and gets a diff that looks like a host bug.
    out.sort_by(|a, b| a.1.cmp(&b.1));
    out
}

/// The `name`s alone, sorted — the common case, spelled once.
pub fn read_directory_names(dir: &Path) -> Vec<String> {
    read_directory(dir)
        .into_iter()
        .map(|(_, name)| name)
        .collect()
}

fn entry_pair(value: &Value) -> (String, String) {
    let Value::Object(object) = value else {
        panic!("expected a directory-entry record, got {value:?}");
    };
    let object = object.lock().unwrap();
    let field = |key: &str| match object.properties.get(key) {
        Some(Value::String(text)) => text.to_string(),
        other => panic!("directory-entry.{key} must be a string, got {other:?}"),
    };
    (field("type"), field("name"))
}

fn bytes_to_vec(value: &Value) -> Vec<u8> {
    let Value::Object(object) = value else {
        panic!("expected a byte array, got {value:?}");
    };
    let object = object.lock().unwrap();
    let ObjectKind::Array(bytes) = &object.kind else {
        panic!("expected an array, got {:?}", object.kind);
    };
    bytes
        .iter()
        .map(|value| match value {
            Value::I32(byte) => *byte as u8,
            Value::F64(byte) => *byte as u8,
            other => panic!("expected a byte, got {other:?}"),
        })
        .collect()
}
