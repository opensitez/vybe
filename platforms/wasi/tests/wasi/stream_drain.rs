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
use vybe_runtime::chunk::ActiveDataSegment;

#[test]
fn canonical_allocator_handles_multi_page_buffers() {
    for size in [65_536, 65_537, 256 * 1024 + 17] {
        let mut chunk = Chunk::new("<canonical-large-allocation>");
        chunk.memory_min_pages.push(1);
        let n = chunk.alloc_scratch(1);
        chunk.emit_i32_const(size, 0);
        chunk.emit_op_u16(Op::LOCAL_SET, n, 0);
        let ptr = vybe_compiler::primitives::canon_marshal::emit_alloc(&mut chunk, 0, n);
        chunk.emit_op_u16(Op::LOCAL_GET, ptr, 0);
        chunk.emit_op_u16(Op::LOCAL_GET, n, 0);
        chunk.emit_op(Op::I32_ADD, 0);
        chunk.emit_i32_const(1, 0);
        chunk.emit_op(Op::I32_SUB, 0);
        chunk.emit_i32_const(0x7f, 0);
        chunk.emit_op(Op::I32_STORE8, 0);
        chunk.emit_op_u16(Op::LOCAL_GET, ptr, 0);
        chunk.emit_op(Op::RETURN, 0);

        let mut vm = VM::new();
        register_platforms(&mut vm, &Capabilities::all());
        let base = vm
            .run(vec![chunk])
            .expect("large canonical allocation")
            .as_i32() as usize;
        let end = base + size as usize;
        assert!(
            end <= vm.memory.len(),
            "{size} bytes at {base}, memory {}",
            vm.memory.len()
        );
        assert_eq!(vm.memory.len(), end.div_ceil(65_536) * 65_536);
        assert_eq!(vm.memory.read_array::<1>(end - 1).unwrap(), [0x7f]);
    }
}

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

/// Construct the same WASI stream without draining it through the canonical ABI.
/// Used only to separate host production from guest consumption in benchmarks.
pub fn produce_read_via_stream(dir: &Path, name: &str) {
    let mut chunk = Chunk::new("<wasi-fs-stream-produce>");
    let preopens = chunk.add_import("wasi:filesystem/preopens", "get-directories");
    let at = chunk.add_import("ecma:array", "at");
    let open_at = chunk.add_import("wasi:filesystem/types", "[method]descriptor.open-at");
    let read_via = chunk.add_import("wasi:filesystem/types", "[method]descriptor.read-via-stream");

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
    chunk.emit_f64_const(0.0, 0);
    chunk.emit_call(read_via, 2, 0);
    chunk.emit_op(Op::RETURN, 0);

    let mut vm = VM::new();
    register_platforms(&mut vm, &Capabilities::all());
    let result = vm.run(vec![chunk]).expect("VM run failed");
    assert!(matches!(result, Value::Object(_)));
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

#[test]
#[ignore = "local canonical byte-array append diagnostic"]
fn bench_canonical_byte_array_append_only() {
    let data: Vec<u8> = (0..256 * 1024).map(|index| (index % 251) as u8).collect();
    let mut chunk = Chunk::new("<canonical-byte-array-append-only>");
    chunk.memory_min_pages.push(4);
    chunk.data_segments.push(data.clone());
    chunk.active_data_segments.push(ActiveDataSegment {
        memory_index: 0,
        offset: 0,
        data_index: 0,
    });
    vybe_compiler::primitives::canon_marshal::emit_new_bytes(&mut chunk, 0);
    let array = chunk.alloc_scratch(1);
    let ptr = chunk.alloc_scratch(1);
    let len = chunk.alloc_scratch(1);
    chunk.emit_op_u16(Op::LOCAL_SET, array, 0);
    chunk.emit_i32_const(0, 0);
    chunk.emit_op_u16(Op::LOCAL_SET, ptr, 0);
    chunk.emit_i32_const(data.len() as i32, 0);
    chunk.emit_op_u16(Op::LOCAL_SET, len, 0);
    vybe_compiler::primitives::canon_marshal::emit_append_bytes(&mut chunk, 0, array, ptr, len);
    chunk.emit_op_u16(Op::LOCAL_GET, array, 0);
    chunk.emit_op(Op::RETURN, 0);

    let mut vm = VM::new();
    register_platforms(&mut vm, &Capabilities::all());
    let started = std::time::Instant::now();
    let result = vm.run(vec![chunk]).expect("append-only module");
    let elapsed = started.elapsed();
    assert_eq!(bytes_to_vec(&result), data);
    println!("canonical 256 KiB append only: {elapsed:?}");
}

#[test]
#[ignore = "local canonical small-chunk append diagnostic"]
fn bench_canonical_byte_array_append_small_chunks() {
    let data: Vec<u8> = (0..256 * 1024).map(|index| (index % 251) as u8).collect();
    let mut chunk = Chunk::new("<canonical-byte-array-small-chunks>");
    chunk.memory_min_pages.push(4);
    chunk.data_segments.push(data.clone());
    chunk.active_data_segments.push(ActiveDataSegment {
        memory_index: 0,
        offset: 0,
        data_index: 0,
    });
    vybe_compiler::primitives::canon_marshal::emit_new_bytes(&mut chunk, 0);
    let array = chunk.alloc_scratch(1);
    let ptr = chunk.alloc_scratch(1);
    let len = chunk.alloc_scratch(1);
    chunk.emit_op_u16(Op::LOCAL_SET, array, 0);
    chunk.emit_i32_const(0, 0);
    chunk.emit_op_u16(Op::LOCAL_SET, ptr, 0);
    chunk.emit_i32_const(64, 0);
    chunk.emit_op_u16(Op::LOCAL_SET, len, 0);
    let done = chunk.emit_block(0);
    let (body, _) = chunk.emit_loop_s(0);
    chunk.emit_op_u16(Op::LOCAL_GET, ptr, 0);
    chunk.emit_i32_const(data.len() as i32, 0);
    chunk.emit_op(Op::I32_GE_U, 0);
    chunk.emit_br_if(1, 0);
    vybe_compiler::primitives::canon_marshal::emit_append_bytes(&mut chunk, 0, array, ptr, len);
    chunk.emit_op_u16(Op::LOCAL_GET, ptr, 0);
    chunk.emit_i32_const(64, 0);
    chunk.emit_op(Op::I32_ADD, 0);
    chunk.emit_op_u16(Op::LOCAL_SET, ptr, 0);
    chunk.emit_br(0, 0);
    chunk.emit_end(0);
    chunk.patch_loop(body);
    chunk.emit_end(0);
    chunk.patch_block(done);
    chunk.emit_op_u16(Op::LOCAL_GET, array, 0);
    chunk.emit_op(Op::RETURN, 0);

    let mut samples = Vec::with_capacity(5);
    for _ in 0..5 {
        let mut vm = VM::new();
        register_platforms(&mut vm, &Capabilities::all());
        let started = std::time::Instant::now();
        let result = vm.run(vec![chunk.clone()]).expect("small-chunk append");
        samples.push(started.elapsed());
        assert_eq!(bytes_to_vec(&result), data);
    }
    samples.sort_unstable();
    println!("canonical 256 KiB append in 64-byte chunks median: {:?}; samples: {samples:?}", samples[2]);
}

#[test]
fn canonical_byte_append_preserves_batch_boundaries() {
    for size in [0, 1, 63, 64, 65, 127, 128, 129, 253, 254, 255, 256, 257] {
        let data: Vec<u8> = (0..size).map(|index| (index % 251) as u8).collect();
        let mut chunk = Chunk::new("<canonical-byte-append-boundary>");
        chunk.memory_min_pages.push(1);
        chunk.data_segments.push(data.clone());
        chunk.active_data_segments.push(ActiveDataSegment {
            memory_index: 0,
            offset: 0,
            data_index: 0,
        });
        vybe_compiler::primitives::canon_marshal::emit_new_bytes(&mut chunk, 0);
        let array = chunk.alloc_scratch(1);
        let ptr = chunk.alloc_scratch(1);
        let len = chunk.alloc_scratch(1);
        chunk.emit_op_u16(Op::LOCAL_SET, array, 0);
        chunk.emit_i32_const(0, 0);
        chunk.emit_op_u16(Op::LOCAL_SET, ptr, 0);
        chunk.emit_i32_const(size as i32, 0);
        chunk.emit_op_u16(Op::LOCAL_SET, len, 0);
        vybe_compiler::primitives::canon_marshal::emit_append_bytes(&mut chunk, 0, array, ptr, len);
        chunk.emit_op_u16(Op::LOCAL_GET, array, 0);
        chunk.emit_op(Op::RETURN, 0);

        let mut vm = VM::new();
        register_platforms(&mut vm, &Capabilities::all());
        let result = vm.run(vec![chunk]).expect("append-boundary module");
        assert_eq!(bytes_to_vec(&result), data, "size {size}");
    }
}

#[test]
fn canonical_utf8_store_uses_bulk_copy_without_changing_bytes() {
    for text in [
        String::new(),
        "a".into(),
        "é🙂".into(),
        format!("{}é", "x".repeat(63)),
    ] {
        let mut chunk = Chunk::new("<canonical-utf8-store>");
        chunk.memory_min_pages.push(1);
        let (ptr, len) = vybe_compiler::primitives::canon_marshal::emit_store_utf8(
            &mut chunk,
            0,
            |chunk| chunk.emit_string_const(&text, 0),
        );
        chunk.emit_op_u16(Op::LOCAL_GET, ptr, 0);
        chunk.emit_op_u16(Op::LOCAL_GET, len, 0);
        chunk.emit_op(Op::I32_ADD, 0);
        chunk.emit_op(Op::RETURN, 0);

        let mut vm = VM::new();
        register_platforms(&mut vm, &Capabilities::all());
        vm.record_runtime_perf(true);
        let end = vm.run(vec![chunk]).expect("UTF-8 store").as_i32() as usize;
        let mut bytes = vec![0; text.len()];
        assert_eq!(vm.memory.read_bytes(end - text.len(), &mut bytes), text.len());
        assert_eq!(bytes, text.as_bytes());
        if !text.is_empty() {
            assert!(
                vm.take_runtime_perf().unwrap().super_managed_to_linear > 0,
                "encoded byte array did not take the bulk-copy path"
            );
        }
    }
}

#[test]
fn canonical_stream_write_then_read_preserves_binary_bytes() {
    for bytes in [
        vec![],
        vec![0],
        vec![0, 1, 127, 128, 254, 255, 2],
        vec![0, 1, 127, 128, 254, 255, 2, 3],
        vec![0, 1, 127, 128, 254, 255, 2, 3, 4],
    ] {
        let mut chunk = Chunk::new("<canonical-byte-stream-roundtrip>");
        vybe_compiler::primitives::canon_marshal::emit_new_bytes(&mut chunk, 0);
        let array = chunk.alloc_scratch(1);
        chunk.emit_op_u16(Op::LOCAL_SET, array, 0);
        let push = chunk.add_import("ecma:array", "push");
        for &byte in &bytes {
            chunk.emit_op_u16(Op::LOCAL_GET, array, 0);
            chunk.emit_i32_const(byte as i32, 0);
            chunk.emit_call(push, 2, 0);
            chunk.emit_op(Op::DROP, 0);
        }
        chunk.emit_op_u16(Op::LOCAL_GET, array, 0);
        vybe_compiler::primitives::io::emit_bytes_to_stream(&mut chunk, 0);
        vybe_compiler::primitives::io::emit_read_stream_to_bytes(&mut chunk, 0);
        chunk.emit_op(Op::RETURN, 0);

        let mut vm = VM::new();
        register_platforms(&mut vm, &Capabilities::all());
        let result = vm.run(vec![chunk]).expect("canonical stream round-trip");
        assert_eq!(bytes_to_vec(&result), bytes);
    }
}
