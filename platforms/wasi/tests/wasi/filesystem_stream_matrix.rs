//! `read-via-stream` over the offset dimension, 0.3.1.
//!
//! Every case here goes through `stream_drain`, which builds ONE chunk doing
//! open-at → read-via-stream → `.at(0)` → `canon stream.read`. The
//! one-host-call-per-VM `invoke` helper the rest of these files use cannot
//! serve: a stream end is an index into that VM's handle table, so handing the
//! tuple out and calling back in yields a handle from a VM that is gone.

use std::path::PathBuf;
use std::sync::atomic::{AtomicU64, Ordering};
use vybe_compiler::primitives::platforms::register_platforms;
use vybe_runtime::capabilities::Capabilities;
use vybe_runtime::value::Object;
use vybe_runtime::{Chunk, Op, VM, Value};

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

/// The offset dimension of `read-via-stream`, drained the 0.3.1 way.
///
/// The LENGTH dimension is gone, and its absence is the point: 0.2's
/// `input-stream.read(n)` took a byte count, while 0.3.1's stream carries
/// everything from `offset` and the guest's drain loop decides how much of it
/// to keep. These cases used to assert `(offset, length) -> slice`; there is
/// no length to pass any more, so each one asserts the whole suffix.
macro_rules! read_case {
    ($name:ident, $offset:expr, $expected:expr) => {
        #[test]
        fn $name() {
            let dir = scratch_dir(stringify!($name));
            std::fs::write(dir.join("payload.bin"), b"abcdef").unwrap();
            let bytes = crate::stream_drain::read_via_stream(&dir, "payload.bin", $offset);
            assert_eq!(bytes, $expected);
        }
    };
}

read_case!(read_from_offset_zero_returns_whole_file, 0.0, b"abcdef");
read_case!(read_from_offset_one_returns_suffix, 1.0, b"bcdef");
read_case!(read_from_offset_four_returns_remaining_suffix, 4.0, b"ef");

// Reading AT the end and PAST it are different conditions that happen to
// share an answer here: both drain empty. Kept as separate cases because a
// future short-read model has to keep them distinguishable.
read_case!(read_from_exact_end_returns_empty_array, 6.0, b"");
read_case!(read_from_beyond_end_returns_empty_array, 12.0, b"");

fn canonical_write_bytes(path: &std::path::Path, payload: &[u8], append: bool) {
    let mut chunk = Chunk::new("<wasi-fs-stream-write>");
    let preopens = chunk.add_import("wasi:filesystem/preopens", "get-directories");
    let at = chunk.add_import("ecma:array", "at");
    let open_at = chunk.add_import("wasi:filesystem/types", "[method]descriptor.open-at");
    let write_via = chunk.add_import(
        "wasi:filesystem/types",
        if append {
            "[method]descriptor.append-via-stream"
        } else {
            "[method]descriptor.write-via-stream"
        },
    );
    let descriptor = chunk.alloc_scratch(1);
    let readable = chunk.alloc_scratch(1);
    let writable = chunk.alloc_scratch(1);
    let payload_name = chunk.intern_string_constant("__wasi_fs_write_payload");

    chunk.emit_call(preopens, 0, 0);
    chunk.emit_i32_const(0, 0);
    chunk.emit_call(at, 2, 0);
    chunk.emit_i32_const(0, 0);
    chunk.emit_call(at, 2, 0);
    chunk.emit_i32_const(0, 0);
    chunk.emit_string_const(&path.display().to_string(), 0);
    chunk.emit_i32_const(0, 0);
    chunk.emit_i32_const(2, 0);
    chunk.emit_call(open_at, 5, 0);
    chunk.emit_op_u16(Op::LOCAL_SET, descriptor, 0);

    vybe_compiler::primitives::io::emit_payload_via_stream(
        &mut chunk,
        readable,
        writable,
        0,
        vybe_compiler::primitives::io::Payload::Bytes,
        |chunk| chunk.emit_op_u32(Op::GLOBAL_GET, payload_name, 0),
        |chunk, readable| {
            chunk.emit_op_u16(Op::LOCAL_GET, descriptor, 0);
            chunk.emit_op_u16(Op::LOCAL_GET, readable, 0);
            if !append {
                chunk.emit_f64_const(6.0, 0);
            }
            chunk.emit_call(write_via, if append { 2 } else { 3 }, 0);
            chunk.emit_op(Op::DROP, 0);
        },
    );
    chunk.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, 0);
    chunk.emit_op(Op::RETURN, 0);

    let mut vm = VM::new();
    register_platforms(&mut vm, &Capabilities::all());
    vm.set_global_owned(
        "__wasi_fs_write_payload",
        Value::Object(vybe_runtime::heap::alloc(Object::new_array(
            payload.iter().map(|byte| Value::I32(*byte as i32)).collect(),
        ))),
    );
    vm.run(vec![chunk]).expect("canonical file write");
}

#[test]
fn canonical_write_preserves_binary_bytes_across_chunk_boundary() {
    let dir = scratch_dir("write_binary");
    let path = dir.join("payload.bin");
    std::fs::write(&path, b"prefix").unwrap();
    let payload: Vec<u8> = (0..131_073).map(|index| (index % 251) as u8).collect();
    canonical_write_bytes(&path, &payload, false);
    let written = std::fs::read(&path).unwrap();
    assert_eq!(&written[..6], b"prefix");
    assert_eq!(&written[6..], payload);
    std::fs::remove_dir_all(dir).unwrap();
}

#[test]
fn canonical_append_preserves_binary_bytes_across_chunk_boundary() {
    let dir = scratch_dir("append_binary");
    let path = dir.join("payload.bin");
    std::fs::write(&path, b"prefix").unwrap();
    let payload: Vec<u8> = (0..131_073).map(|index| (index % 251) as u8).collect();
    canonical_write_bytes(&path, &payload, true);
    let written = std::fs::read(&path).unwrap();
    assert_eq!(&written[..6], b"prefix");
    assert_eq!(&written[6..], payload);
    std::fs::remove_dir_all(dir).unwrap();
}

#[test]
#[ignore = "performance measurement; run explicitly with --ignored --nocapture"]
fn bench_bulk_write_via_canonical_stream() {
    let dir = scratch_dir("write_perf");
    let path = dir.join("payload.bin");
    let payload: Vec<u8> = (0..256 * 1024).map(|index| (index % 251) as u8).collect();
    let mut samples = Vec::with_capacity(5);
    for _ in 0..5 {
        std::fs::write(&path, b"prefix").unwrap();
        let started = std::time::Instant::now();
        canonical_write_bytes(&path, &payload, false);
        samples.push(started.elapsed());
        let written = std::fs::read(&path).unwrap();
        assert_eq!(&written[..6], b"prefix");
        assert_eq!(&written[6..], payload);
    }
    samples.sort_unstable();
    println!(
        "canonical 256 KiB file write median: {:?}; samples: {samples:?}",
        samples[2]
    );
    std::fs::remove_dir_all(dir).unwrap();
}

fn byte_array_store_chunk() -> Chunk {
    let mut chunk = Chunk::new("<canonical-byte-store>");
    let payload_name = chunk.intern_string_constant("__wasi_fs_write_payload");
    let (ptr, _len) = vybe_compiler::primitives::canon_marshal::emit_store_byte_array(
        &mut chunk,
        0,
        |chunk| chunk.emit_op_u32(Op::GLOBAL_GET, payload_name, 0),
    );
    chunk.emit_op_u16(Op::LOCAL_GET, ptr, 0);
    chunk.emit_op(Op::RETURN, 0);
    chunk
}

#[test]
fn canonical_byte_store_preserves_numeric_types_and_traps() {
    let run = |values: Vec<Value>| {
        let mut vm = VM::new();
        register_platforms(&mut vm, &Capabilities::all());
        vm.set_global_owned(
            "__wasi_fs_write_payload",
            Value::Object(vybe_runtime::heap::alloc(Object::new_array(values))),
        );
        let result = vm.run(vec![byte_array_store_chunk()]);
        (vm, result)
    };

    let (vm, result) = run(vec![Value::I32(0), Value::I32(255), Value::F64(42.0)]);
    let ptr = result.expect("integer-valued numbers").as_i32() as usize;
    let mut actual = [0; 3];
    assert_eq!(vm.memory.read_bytes(ptr, &mut actual), actual.len());
    assert_eq!(actual, [0, 255, 42]);

    for invalid in [Value::F64(1.5), Value::F64(-0.0), Value::String("7".into())] {
        let (_, result) = run(vec![Value::I32(7), invalid]);
        assert!(result.is_err(), "non-i32 element must retain toI32 trap");
    }
}

#[test]
fn canonical_byte_store_bulk_path_requires_plain_i32_values() {
    let run = |values: Vec<Value>| {
        let mut vm = VM::new();
        register_platforms(&mut vm, &Capabilities::all());
        vm.set_global_owned(
            "__wasi_fs_write_payload",
            Value::Object(vybe_runtime::heap::alloc(Object::new_array(values))),
        );
        vm.record_runtime_perf(true);
        let result = vm.run(vec![byte_array_store_chunk()]);
        let counters = vm.take_runtime_perf().expect("runtime counters");
        (result, counters)
    };

    let (result, counters) = run(vec![Value::I32(0), Value::I32(255)]);
    assert!(result.is_ok());
    assert_eq!(counters.super_managed_to_linear, 1);
    assert_eq!(counters.call_target_js_builtin, 0);

    let (result, counters) = run(vec![Value::I32(0), Value::F64(42.0)]);
    assert!(result.is_ok());
    assert_eq!(counters.super_managed_to_linear, 0);
    assert!(counters.call_target_js_builtin > 0);

    let (result, counters) = run(vec![Value::I32(0), Value::F64(1.5)]);
    assert!(result.is_err());
    assert_eq!(counters.super_managed_to_linear, 0);
}

#[test]
#[ignore = "performance measurement; run explicitly with --ignored --nocapture"]
fn bench_canonical_byte_array_store_only() {
    let payload: Vec<u8> = (0..256 * 1024).map(|index| (index % 251) as u8).collect();
    let chunk = byte_array_store_chunk();

    let mut samples = Vec::with_capacity(5);
    for _ in 0..5 {
        let mut vm = VM::new();
        register_platforms(&mut vm, &Capabilities::all());
        vm.set_global_owned(
            "__wasi_fs_write_payload",
            Value::Object(vybe_runtime::heap::alloc(Object::new_array(
                payload.iter().map(|byte| Value::I32(*byte as i32)).collect(),
            ))),
        );
        let started = std::time::Instant::now();
        let result = vm.run(vec![chunk.clone()]).expect("canonical byte store");
        samples.push(started.elapsed());
        let ptr = result.as_i32() as u32;
        let mut written = vec![0; payload.len()];
        assert_eq!(vm.memory.read_bytes(ptr as usize, &mut written), payload.len());
        assert_eq!(written, payload);
    }
    samples.sort_unstable();
    println!(
        "canonical 256 KiB byte-array store median: {:?}; samples: {samples:?}",
        samples[2]
    );

    let mut vm = VM::new();
    register_platforms(&mut vm, &Capabilities::all());
    vm.set_global_owned(
        "__wasi_fs_write_payload",
        Value::Object(vybe_runtime::heap::alloc(Object::new_array(
            payload.iter().map(|byte| Value::I32(*byte as i32)).collect(),
        ))),
    );
    vm.record_runtime_perf(true);
    vm.run(vec![chunk]).expect("instrumented canonical byte store");
    let counters = vm.take_runtime_perf().unwrap();
    println!(
        "canonical store calls: js_builtin={}, host={}, bulk_copies={}",
        counters.call_target_js_builtin,
        counters.host_call,
        counters.super_managed_to_linear
    );
}

#[test]
fn canonical_stream_preserves_batch_edges_and_binary_bytes() {
    let dir = scratch_dir("batch_edges");
    for len in [0, 1, 7, 8, 9, 15, 16, 17, 63, 64, 65, 127, 128, 129, 191, 192, 193, 383, 384, 385, 65535, 65536, 65537] {
        let data: Vec<u8> = (0..len).map(|i| (i % 251) as u8).collect();
        std::fs::write(dir.join("payload.bin"), &data).unwrap();
        assert_eq!(
            crate::stream_drain::read_via_stream(&dir, "payload.bin", 0.0),
            data,
            "length {len}"
        );
    }
    std::fs::remove_dir_all(dir).unwrap();
}

#[test]
fn canonical_stream_preserves_utf8_split_across_read_boundary() {
    let dir = scratch_dir("utf8_boundary");
    let mut data = vec![b'a'; 65535];
    data.extend_from_slice("é".as_bytes());
    data.push(b'z');
    std::fs::write(dir.join("payload.txt"), &data).unwrap();
    assert_eq!(
        crate::stream_drain::read_via_stream(&dir, "payload.txt", 0.0),
        data
    );
    std::fs::remove_dir_all(dir).unwrap();
}

/// Each call gets its OWN stream, positioned by its own offset.
///
/// The 0.2 shape had one `input-stream` resource whose cursor advanced across
/// reads, so "sequential reads advance the position" was a property of the
/// resource. 0.3.1 has no such resource: `read-via-stream` behaves like
/// `pread`, so position is the argument and successive reads are independent.
/// Asserting the old property against the new function would assert a cursor
/// that no longer exists.
#[test]
fn each_read_is_positioned_independently() {
    let dir = scratch_dir("independent_positions");
    std::fs::write(dir.join("payload.bin"), b"abcdef").unwrap();
    assert_eq!(
        crate::stream_drain::read_via_stream(&dir, "payload.bin", 0.0),
        b"abcdef"
    );
    assert_eq!(
        crate::stream_drain::read_via_stream(&dir, "payload.bin", 3.0),
        b"def"
    );
    assert_eq!(
        crate::stream_drain::read_via_stream(&dir, "payload.bin", 0.0),
        b"abcdef",
        "a second read from 0 sees the whole file again — there is no cursor to have moved"
    );
}

/// `tuple<stream<u8>, future<result<_, error-code>>>`: element 1 resolves with
/// success on a clean read. `Null` is the `ok` case; an `error-code` arrives as
/// a string.
#[test]
fn clean_read_resolves_the_outcome_future_with_ok() {
    let dir = scratch_dir("outcome_ok");
    std::fs::write(dir.join("payload.bin"), b"abcdef").unwrap();
    let outcome = crate::stream_drain::read_via_stream_outcome(&dir, "payload.bin", 0.0);
    assert!(
        !matches!(outcome, vybe_runtime::value::Value::String(_)),
        "a clean read must not resolve the future with an error-code, got {outcome:?}"
    );
}

#[test]
#[ignore = "performance measurement; run explicitly with --ignored --nocapture"]
fn bench_bulk_read_via_canonical_stream() {
    let dir = scratch_dir("bulk_perf");
    let data: Vec<u8> = (0..256 * 1024).map(|i| (i % 251) as u8).collect();
    std::fs::write(dir.join("payload.bin"), &data).unwrap();

    let started = std::time::Instant::now();
    let actual = crate::stream_drain::read_via_stream(&dir, "payload.bin", 0.0);
    let elapsed = started.elapsed();
    assert_eq!(actual, data);
    println!("canonical 256 KiB file read: {elapsed:?}");

    std::fs::remove_dir_all(dir).unwrap();
}

#[test]
#[ignore = "performance measurement; run explicitly with --ignored --nocapture"]
fn bench_text_read_via_canonical_stream() {
    let dir = scratch_dir("text_perf");
    let data = b"The quick brown fox jumps over the lazy dog.\n".repeat(6_000);
    std::fs::write(dir.join("payload.txt"), &data).unwrap();

    let started = std::time::Instant::now();
    let actual = crate::stream_drain::read_via_stream(&dir, "payload.txt", 0.0);
    let elapsed = started.elapsed();
    assert_eq!(actual, data);
    println!("canonical {} KiB text file read: {elapsed:?}", data.len() / 1024);

    std::fs::remove_dir_all(dir).unwrap();
}

#[test]
#[ignore = "performance measurement; run explicitly with --ignored --nocapture"]
fn bench_bulk_stream_production_only() {
    let dir = scratch_dir("produce_perf");
    let data: Vec<u8> = (0..256 * 1024).map(|i| (i % 251) as u8).collect();
    std::fs::write(dir.join("payload.bin"), data).unwrap();

    let started = std::time::Instant::now();
    crate::stream_drain::produce_read_via_stream(&dir, "payload.bin");
    println!("produce 256 KiB stream, no canonical drain: {:?}", started.elapsed());

    std::fs::remove_dir_all(dir).unwrap();
}

#[test]
#[ignore = "performance measurement; run explicitly with --ignored --nocapture"]
fn bench_text_stream_production_only() {
    let dir = scratch_dir("produce_text_perf");
    let data = b"The quick brown fox jumps over the lazy dog.\n".repeat(6_000);
    std::fs::write(dir.join("payload.txt"), data).unwrap();

    let started = std::time::Instant::now();
    crate::stream_drain::produce_read_via_stream(&dir, "payload.txt");
    println!("produce 263 KiB text stream, no canonical drain: {:?}", started.elapsed());

    std::fs::remove_dir_all(dir).unwrap();
}
