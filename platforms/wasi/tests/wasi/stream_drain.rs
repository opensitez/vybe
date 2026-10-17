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
use vybe_runtime::chunk::ActiveDataSegment;
use vybe_runtime::value::{ObjectKind, Value};
use vybe_runtime::{Chunk, Op, VM};

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
    let chunk = build_read_via_stream_drain(dir, name, offset);
    let mut vm = VM::new();
    register_platforms(&mut vm, &Capabilities::all());
    let result = vm.run(vec![chunk]).expect("VM run failed");
    bytes_to_vec(&result)
}

pub(crate) fn build_read_via_stream_drain(dir: &Path, name: &str, offset: f64) -> Chunk {
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
    chunk
}

/// Construct the same WASI stream without draining it through the canonical ABI.
/// Used only to separate host production from guest consumption in benchmarks.
pub fn produce_read_via_stream(dir: &Path, name: &str) {
    let chunk = build_read_via_stream_producer(dir, name);
    let mut vm = VM::new();
    register_platforms(&mut vm, &Capabilities::all());
    let result = vm.run(vec![chunk]).expect("VM run failed");
    assert!(matches!(result, Value::Object(_)));
}

pub(crate) fn build_read_via_stream_producer(dir: &Path, name: &str) -> Chunk {
    let mut chunk = Chunk::new("<wasi-fs-stream-produce>");
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
    chunk.emit_f64_const(0.0, 0);
    chunk.emit_call(read_via, 2, 0);
    chunk.emit_op(Op::RETURN, 0);
    chunk
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

    let mut chunks = vec![chunk];
    let (mut definitions, mut functions, mut values) = (Vec::new(), Vec::new(), Vec::new());
    vybe_compiler::primitives::canon::resolve_emitter_imports(
        &mut chunks, &mut definitions, &mut functions, &mut values).unwrap();
    vybe_compiler::primitives::canon::install_canon_section(&mut chunks, definitions);
    vybe_compiler::primitives::canon::install_type_space(
        &mut chunks, functions, values, Vec::new());

    let mut vm = VM::new();
    register_platforms(&mut vm, &Capabilities::all());
    let result = vm.run(chunks).expect("VM run failed");

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
fn canonical_byte_append_bulk_batch_preserves_tail() {
    for size in [
        0, 1, 63, 64, 65, 127, 128, 253, 254, 255, 317, 318, 319, 508, 509, 510, 572,
    ] {
        let data: Vec<u8> = (0..size).map(|i| (i % 251) as u8).collect();
        for custom_property in [false, true] {
            let mut chunk = Chunk::new("<canonical-byte-append-batch>");
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
            if custom_property {
                chunk.emit_op_u16(Op::LOCAL_GET, array, 0);
                chunk.emit_string_const("marker", 0);
                chunk.emit_i32_const(1, 0);
                let set = chunk.add_import("ecma:object", "set");
                chunk.emit_call(set, 3, 0);
                chunk.emit_op(Op::DROP, 0);
            }
            chunk.emit_i32_const(0, 0);
            chunk.emit_op_u16(Op::LOCAL_SET, ptr, 0);
            chunk.emit_i32_const(data.len() as i32, 0);
            chunk.emit_op_u16(Op::LOCAL_SET, len, 0);
            vybe_compiler::primitives::canon_marshal::emit_append_bytes(
                &mut chunk, 0, array, ptr, len,
            );
            chunk.emit_op_u16(Op::LOCAL_GET, array, 0);
            chunk.emit_op(Op::RETURN, 0);

            let mut vm = VM::new();
            register_platforms(&mut vm, &Capabilities::all());
            vm.record_runtime_perf(true);
            let result = vm.run(vec![chunk]).expect("canonical byte append");
            assert_eq!(bytes_to_vec(&result), data);
            let counters = vm.take_runtime_perf().expect("runtime counters");
            let expected = if custom_property {
                0
            } else {
                let remaining = size % 254;
                u64::from(size >= 254) + u64::from(remaining >= 64) + u64::from(remaining % 64 != 0)
            };
            assert_eq!(
                counters.super_linear_to_managed, expected,
                "size={size}, custom_property={custom_property}"
            );
        }
    }
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
    println!(
        "canonical 256 KiB append in 64-byte chunks median: {:?}; samples: {samples:?}",
        samples[2]
    );
}

#[test]
#[ignore = "local canonical append fast-path comparison"]
fn bench_canonical_byte_array_append_guarded_vs_fallback() {
    let data: Vec<u8> = (0..256 * 1024).map(|index| (index % 251) as u8).collect();
    let mut chunks = Vec::new();
    for custom_property in [false, true] {
        let mut chunk = Chunk::new("<canonical-byte-array-append-comparison>");
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
        if custom_property {
            chunk.emit_op_u16(Op::LOCAL_GET, array, 0);
            chunk.emit_string_const("marker", 0);
            chunk.emit_i32_const(1, 0);
            let set = chunk.add_import("ecma:object", "set");
            chunk.emit_call(set, 3, 0);
            chunk.emit_op(Op::DROP, 0);
        }
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
        chunks.push(chunk);
    }

    let mut samples = [Vec::new(), Vec::new()];
    for round in 0..7 {
        for variant in [round % 2, (round + 1) % 2] {
            let mut vm = VM::new();
            register_platforms(&mut vm, &Capabilities::all());
            let chunk = chunks[variant].clone();
            let started = std::time::Instant::now();
            let result = vm.run(vec![chunk]).expect("append comparison module");
            let elapsed = started.elapsed();
            assert_eq!(bytes_to_vec(&result), data);
            if round != 0 {
                samples[variant].push(elapsed);
            }
        }
    }
    for (label, runs) in ["guarded", "fallback"].into_iter().zip(&mut samples) {
        runs.sort_unstable();
        println!(
            "canonical 256 KiB append, 64-byte chunks, {label}: {:?}",
            runs[3]
        );
    }
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
        let (ptr, len) =
            vybe_compiler::primitives::canon_marshal::emit_store_utf8(&mut chunk, 0, |chunk| {
                chunk.emit_string_const(&text, 0)
            });
        chunk.emit_op_u16(Op::LOCAL_GET, ptr, 0);
        chunk.emit_op_u16(Op::LOCAL_GET, len, 0);
        chunk.emit_op(Op::I32_ADD, 0);
        chunk.emit_op(Op::RETURN, 0);

        let mut vm = VM::new();
        register_platforms(&mut vm, &Capabilities::all());
        vm.record_runtime_perf(true);
        let end = vm.run(vec![chunk]).expect("UTF-8 store").as_i32() as usize;
        let mut bytes = vec![0; text.len()];
        assert_eq!(
            vm.memory.read_bytes(end - text.len(), &mut bytes),
            text.len()
        );
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
#[ignore = "local canonical UTF-8 store diagnostic"]
fn bench_canonical_utf8_store_only() {
    let text = "The quick brown fox jumps over the lazy dog.\n".repeat(6_000);
    let mut chunk = Chunk::new("<canonical-utf8-store-bench>");
    chunk.memory_min_pages.push(1);
    let (ptr, len) =
        vybe_compiler::primitives::canon_marshal::emit_store_utf8(&mut chunk, 0, |chunk| {
            chunk.emit_string_const(&text, 0)
        });
    chunk.emit_op_u16(Op::LOCAL_GET, ptr, 0);
    chunk.emit_op_u16(Op::LOCAL_GET, len, 0);
    chunk.emit_op(Op::I32_ADD, 0);
    chunk.emit_op(Op::RETURN, 0);

    let mut samples = Vec::new();
    for round in 0..6 {
        let mut vm = VM::new();
        register_platforms(&mut vm, &Capabilities::all());
        let started = std::time::Instant::now();
        let end = vm.run(vec![chunk.clone()]).expect("UTF-8 store").as_i32() as usize;
        let elapsed = started.elapsed();
        let mut bytes = vec![0; text.len()];
        assert_eq!(
            vm.memory.read_bytes(end - text.len(), &mut bytes),
            bytes.len()
        );
        assert_eq!(bytes, text.as_bytes());
        if round != 0 {
            samples.push(elapsed);
        }
    }
    samples.sort_unstable();
    println!(
        "canonical {} KiB UTF-8 store median: {:?}",
        text.len() / 1024,
        samples[2]
    );
}

#[test]
#[ignore = "local canonical UTF-8 phase diagnostic"]
fn bench_canonical_utf8_store_phases() {
    let text = "The quick brown fox jumps over the lazy dog.\n".repeat(6_000);
    let make_encode_chunk = |convert: bool| {
        let mut chunk = Chunk::new(if convert {
            "<utf8-array-from-phase>"
        } else {
            "<utf8-encode-phase>"
        });
        let encoder_new = chunk.add_import("web:encoding", "encoderNew");
        chunk.emit_call(encoder_new, 0, 0);
        chunk.emit_string_const(&text, 0);
        let encoder_encode = chunk.add_import("web:encoding", "encode");
        chunk.emit_call(encoder_encode, 2, 0);
        if convert {
            let from = chunk.add_import("ecma:array", "from");
            chunk.emit_call(from, 1, 0);
        }
        chunk.emit_op(Op::RETURN, 0);
        chunk
    };
    let encode = make_encode_chunk(false);
    let array_from = make_encode_chunk(true);

    let mut store = Chunk::new("<utf8-store-phase>");
    store.memory_min_pages.push(1);
    let (ptr, len) =
        vybe_compiler::primitives::canon_marshal::emit_store_utf8(&mut store, 0, |chunk| {
            chunk.emit_string_const(&text, 0)
        });
    store.emit_op_u16(Op::LOCAL_GET, ptr, 0);
    store.emit_op_u16(Op::LOCAL_GET, len, 0);
    store.emit_op(Op::I32_ADD, 0);
    store.emit_op(Op::RETURN, 0);

    for (name, chunk) in [
        ("encode", encode),
        ("array-from", array_from),
        ("store", store),
    ] {
        let mut samples = Vec::new();
        for round in 0..8 {
            let mut vm = VM::new();
            register_platforms(&mut vm, &Capabilities::all());
            let started = std::time::Instant::now();
            let result = vm.run(vec![chunk.clone()]).expect("UTF-8 phase");
            let elapsed = started.elapsed();
            if name == "store" {
                let end = result.as_i32() as usize;
                let mut bytes = vec![0; text.len()];
                assert_eq!(
                    vm.memory.read_bytes(end - text.len(), &mut bytes),
                    bytes.len()
                );
                assert_eq!(bytes, text.as_bytes());
            }
            if round != 0 {
                samples.push(elapsed);
            }
        }
        samples.sort_unstable();
        println!(
            "UTF-8 {name} {} KiB median: {:?}",
            text.len() / 1024,
            samples[3]
        );
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

#[test]
fn canonical_drain_preserves_fragmented_byte_streams() {
    for size in [0, 1, 63, 64, 65, 253, 254, 255, 256, 257, 511, 512, 1025] {
        let data: Vec<u8> = (0..size).map(|i| (i % 251) as u8).collect();
        let mut chunk = Chunk::new("<fragmented-canonical-byte-stream>");
        chunk.memory_min_pages.push(1);
        chunk.data_segments.push(data.clone());
        chunk.active_data_segments.push(ActiveDataSegment {
            memory_index: 0,
            offset: 0,
            data_index: 0,
        });

        let pair = chunk.i64_scratch();
        let readable = chunk.alloc_scratch(1);
        let writable = chunk.alloc_scratch(1);
        let stream_new = chunk.add_import("canon", "stream.new");
        let stream_write = chunk.add_import("canon", "stream.write");
        let drop_writable = chunk.add_import("canon", "stream.drop-writable");
        chunk.emit_call(stream_new, 0, 0);
        chunk.emit_op_u16(Op::LOCAL_SET, pair, 0);
        chunk.emit_op_u16(Op::LOCAL_GET, pair, 0);
        chunk.emit_op(Op::I32_WRAP_I64, 0);
        chunk.emit_op_u16(Op::LOCAL_SET, readable, 0);
        chunk.emit_op_u16(Op::LOCAL_GET, pair, 0);
        chunk.emit_i64_const(32, 0);
        chunk.emit_op(Op::I64_SHR_U, 0);
        chunk.emit_op(Op::I32_WRAP_I64, 0);
        chunk.emit_op_u16(Op::LOCAL_SET, writable, 0);

        for offset in (0..size).step_by(64) {
            chunk.emit_op_u16(Op::LOCAL_GET, writable, 0);
            chunk.emit_i32_const(offset as i32, 0);
            chunk.emit_i32_const((size - offset).min(64) as i32, 0);
            chunk.emit_call(stream_write, 3, 0);
            chunk.emit_op(Op::DROP, 0);
        }
        chunk.emit_op_u16(Op::LOCAL_GET, writable, 0);
        chunk.emit_call(drop_writable, 1, 0);
        chunk.emit_op_u16(Op::LOCAL_GET, readable, 0);
        vybe_compiler::primitives::io::emit_read_stream_to_bytes(&mut chunk, 0);
        chunk.emit_op(Op::RETURN, 0);

        let mut vm = VM::new();
        register_platforms(&mut vm, &Capabilities::all());
        let result = vm.run(vec![chunk]).expect("fragmented canonical stream");
        assert_eq!(bytes_to_vec(&result), data, "size {size}");
    }
}

#[test]
fn canonical_byte_append_assembles_short_chunks() {
    for size in [0, 1, 64, 253, 254, 255, 257, 511, 512, 1025] {
        let data: Vec<u8> = (0..size).map(|i| (i % 251) as u8).collect();
        let mut chunk = Chunk::new("<canonical-short-byte-append>");
        chunk.memory_min_pages.push(1);
        chunk.data_segments.push(data.clone());
        chunk.active_data_segments.push(ActiveDataSegment {
            memory_index: 0,
            offset: 0,
            data_index: 0,
        });

        vybe_compiler::primitives::canon_marshal::emit_new_bytes(&mut chunk, 0);
        let output = chunk.alloc_scratch(1);
        let source = chunk.alloc_scratch(1);
        let count = chunk.alloc_scratch(1);
        chunk.emit_op_u16(Op::LOCAL_SET, output, 0);
        for offset in (0..size).step_by(64) {
            chunk.emit_i32_const(offset as i32, 0);
            chunk.emit_op_u16(Op::LOCAL_SET, source, 0);
            chunk.emit_i32_const((size - offset).min(64) as i32, 0);
            chunk.emit_op_u16(Op::LOCAL_SET, count, 0);
            vybe_compiler::primitives::canon_marshal::emit_append_bytes(
                &mut chunk, 0, output, source, count,
            );
        }
        chunk.emit_op_u16(Op::LOCAL_GET, output, 0);
        chunk.emit_op(Op::RETURN, 0);

        let mut vm = VM::new();
        register_platforms(&mut vm, &Capabilities::all());
        let result = vm.run(vec![chunk]).expect("short byte append");
        assert_eq!(bytes_to_vec(&result), data, "size {size}");
    }
}

#[test]
#[ignore = "local short-read byte assembly benchmark"]
fn bench_canonical_byte_append_chunk_sizes() {
    const SIZE: usize = 256 * 1024;
    let chunk_bytes = std::env::var("VYBE_BENCH_CHUNK_BYTES")
        .map(|value| value.parse::<usize>().expect("numeric benchmark chunk size"))
        .unwrap_or(64);
    assert!(chunk_bytes > 0 && SIZE % chunk_bytes == 0);
    let data: Vec<u8> = (0..SIZE).map(|i| (i % 251) as u8).collect();
    let mut results = [Vec::new()];

    for samples in &mut results {
        let mut chunk = Chunk::new("<canonical-short-read-bench>");
        chunk.memory_min_pages.push(4);
        chunk.data_segments.push(data.clone());
        chunk.active_data_segments.push(ActiveDataSegment {
            memory_index: 0,
            offset: 0,
            data_index: 0,
        });
        vybe_compiler::primitives::canon_marshal::emit_new_bytes(&mut chunk, 0);
        let output = chunk.alloc_scratch(1);
        let source = chunk.alloc_scratch(1);
        let count = chunk.alloc_scratch(1);
        chunk.emit_op_u16(Op::LOCAL_SET, output, 0);
        chunk.emit_i32_const(0, 0);
        chunk.emit_op_u16(Op::LOCAL_SET, source, 0);
        chunk.emit_i32_const(chunk_bytes as i32, 0);
        chunk.emit_op_u16(Op::LOCAL_SET, count, 0);

        let done = chunk.emit_block(0);
        let (loop_start, _) = chunk.emit_loop_s(0);
        chunk.emit_op_u16(Op::LOCAL_GET, source, 0);
        chunk.emit_i32_const(SIZE as i32, 0);
        chunk.emit_op(Op::I32_GE_U, 0);
        chunk.emit_br_if(1, 0);
        vybe_compiler::primitives::canon_marshal::emit_append_bytes(
            &mut chunk, 0, output, source, count,
        );
        chunk.emit_op_u16(Op::LOCAL_GET, source, 0);
        chunk.emit_i32_const(chunk_bytes as i32, 0);
        chunk.emit_op(Op::I32_ADD, 0);
        chunk.emit_op_u16(Op::LOCAL_SET, source, 0);
        chunk.emit_br(0, 0);
        chunk.emit_end(0);
        chunk.patch_loop(loop_start);
        chunk.emit_end(0);
        chunk.patch_block(done);
        chunk.emit_op_u16(Op::LOCAL_GET, output, 0);
        chunk.emit_op(Op::RETURN, 0);

        for _ in 0..5 {
            let mut vm = VM::new();
            register_platforms(&mut vm, &Capabilities::all());
            let input = chunk.clone();
            let start = std::time::Instant::now();
            let result = vm.run(vec![input]).expect("short-read benchmark");
            samples.push(start.elapsed());
            assert_eq!(bytes_to_vec(&result), data);
        }
        samples.sort_unstable();
    }
    println!(
        "256 KiB / {chunk_bytes}-byte chunks: direct {:?}",
        results[0][2]
    );
}
