//! Behaviour tests for `wasi:filesystem` host imports — real WASI
//! 0.2.8 surface (descriptor handles, capability-based paths).
//!
//! Reference WIT: `proposals/wasi-filesystem/wit/{types,preopens}.wit`.
//!
//! Function names match the canonical-ABI shape so a Component-Model
//! runtime loading Vybe-emitted `.wasm` resolves them correctly:
//!   - `[method]descriptor.open-at`
//!   - `[method]descriptor.stat`
//!   - `[method]descriptor.read-via-stream`
//!   - `[method]descriptor.read-directory`
//!   - `[method]directory-entry-stream.read-directory-entry`
//!   - `get-directories` (under `wasi:filesystem/preopens`)

use std::path::PathBuf;
use std::sync::atomic::{AtomicU64, Ordering};
use vybe_bytecode::value::{Object, ObjectKind, Value};
use vybe_bytecode::{Chunk, Op, VM};
use vybe_host::{Capabilities, register_with_capabilities};

// ── Test scaffolding ──────────────────────────────────────────────

fn scratch_dir(label: &str) -> PathBuf {
    static COUNTER: AtomicU64 = AtomicU64::new(0);
    let id = COUNTER.fetch_add(1, Ordering::SeqCst);
    let dir = std::env::temp_dir().join(format!(
        "vybe-wasi-fs-test-{}-{}-{}",
        std::process::id(),
        label,
        id
    ));
    let _ = std::fs::remove_dir_all(&dir);
    std::fs::create_dir_all(&dir).expect("scratch dir mkdir");
    dir
}

fn invoke(module: &str, name: &str, args: Vec<Value>) -> Value {
    let mut chunk = Chunk::new("<wasi-fs-test>");
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
    register_with_capabilities(&mut vm, &Capabilities::all());
    vm.run(vec![chunk]).expect("VM run failed")
}

fn types(name: &str, args: Vec<Value>) -> Value {
    invoke("wasi:filesystem/types", name, args)
}

fn preopens(name: &str, args: Vec<Value>) -> Value {
    invoke("wasi:filesystem/preopens", name, args)
}

fn s(text: &str) -> Value {
    Value::String(std::sync::Arc::from(text))
}

fn prop(value: &Value, key: &str) -> Value {
    if let Value::Object(object) = value {
        let object = object.lock().unwrap();
        if let Some(v) = object.properties.get(key) {
            return v.clone();
        }
    }
    Value::Null
}

fn is_error(value: &Value) -> Option<String> {
    if let Value::Object(object) = value {
        let object = object.lock().unwrap();
        if let Some(Value::String(text)) = object.properties.get("__wasi_error") {
            return Some(text.to_string());
        }
    }
    None
}

/// Open a descriptor for the scratch directory itself — every test
/// starts here. WASI doesn't grant ambient access to absolute paths;
/// in real deployments the host pre-opens directories. This helper
/// stands in for that by opening the scratch root via the
/// `__test_open_root` host fn (registered alongside the WASI
/// surface specifically for tests — production code uses
/// `get-directories`).
fn open_test_root(dir: &PathBuf) -> Value {
    invoke("wasi:filesystem/types", "__test_open_root", vec![s(dir.to_str().unwrap())])
}

// ── preopens.get-directories ──────────────────────────────────────

#[test]
fn get_directories_returns_array_of_pairs() {
    let result = preopens("get-directories", vec![]);
    if let Value::Object(object) = &result {
        let object = object.lock().unwrap();
        if let ObjectKind::Array(elements) = &object.kind {
            for element in elements {
                if let Value::Object(pair) = element {
                    let pair = pair.lock().unwrap();
                    if let ObjectKind::Array(pair_elements) = &pair.kind {
                        assert_eq!(pair_elements.len(), 2, "(descriptor, path) pair shape");
                        // pair[0] is a descriptor object, pair[1] is the path string.
                        assert!(matches!(pair_elements[0], Value::Object(_)), "pair[0] descriptor");
                        assert!(matches!(pair_elements[1], Value::String(_)), "pair[1] path string");
                        return;
                    }
                }
            }
            // Empty preopens is acceptable in capability-restricted
            // environments; the array exists, that's what matters.
            return;
        }
    }
    panic!("get-directories should return an array, got {:?}", result);
}

// ── [method]descriptor.open-at + stat ─────────────────────────────

#[test]
fn open_at_existing_file_returns_descriptor() {
    let dir = scratch_dir("open_at_file");
    let file = dir.join("hello.txt");
    std::fs::write(&file, "hi").unwrap();

    let root = open_test_root(&dir);
    // open-at(parent, path-flags=0, path="hello.txt", open-flags=0, %flags=0)
    let result = types("[method]descriptor.open-at", vec![
        root, Value::I32(0), s("hello.txt"), Value::I32(0), Value::I32(0),
    ]);
    assert!(is_error(&result).is_none(), "expected ok descriptor, got error {:?}", is_error(&result));
    assert!(matches!(result, Value::Object(_)), "open-at returns a descriptor object");
}

#[test]
fn open_at_missing_file_returns_no_entry_error() {
    let dir = scratch_dir("open_at_missing");
    let root = open_test_root(&dir);
    let result = types("[method]descriptor.open-at", vec![
        root, Value::I32(0), s("nope.txt"), Value::I32(0), Value::I32(0),
    ]);
    assert_eq!(is_error(&result).as_deref(), Some("no-entry"));
}

#[test]
fn open_at_with_create_flag_creates_new_file() {
    let dir = scratch_dir("open_at_create");
    let root = open_test_root(&dir);
    // open-flags::create = bit 0 (per WIT: create, directory, exclusive, truncate)
    let result = types("[method]descriptor.open-at", vec![
        root, Value::I32(0), s("new.txt"), Value::I32(1), Value::I32(0),
    ]);
    assert!(is_error(&result).is_none(), "open-at create should succeed");
    assert!(dir.join("new.txt").exists(), "file should be created on disk");
}

#[test]
fn stat_at_returns_descriptor_stat() {
    let dir = scratch_dir("stat_at");
    let file = dir.join("sized.bin");
    std::fs::write(&file, b"abcdef").unwrap();
    let root = open_test_root(&dir);

    // stat-at(path-flags, path)
    let stats = types("[method]descriptor.stat-at", vec![
        root, Value::I32(0), s("sized.bin"),
    ]);
    assert_eq!(prop(&stats, "size"), Value::F64(6.0));
    // descriptor-stat fields per WIT: type, link-count, size, mtime, atime, ctime
    assert_eq!(prop(&stats, "type"), s("regular-file"));
}

#[test]
fn stat_on_open_descriptor() {
    let dir = scratch_dir("stat_descriptor");
    let file = dir.join("d.bin");
    std::fs::write(&file, b"xx").unwrap();
    let root = open_test_root(&dir);

    let descriptor = types("[method]descriptor.open-at", vec![
        root, Value::I32(0), s("d.bin"), Value::I32(0), Value::I32(0),
    ]);
    let stats = types("[method]descriptor.stat", vec![descriptor]);
    assert_eq!(prop(&stats, "size"), Value::F64(2.0));
    assert_eq!(prop(&stats, "type"), s("regular-file"));
}

#[test]
fn get_type_returns_descriptor_type() {
    let dir = scratch_dir("get_type");
    let file = dir.join("t.txt");
    std::fs::write(&file, "").unwrap();
    let root = open_test_root(&dir);

    let descriptor = types("[method]descriptor.open-at", vec![
        root.clone(), Value::I32(0), s("t.txt"), Value::I32(0), Value::I32(0),
    ]);
    assert_eq!(types("[method]descriptor.get-type", vec![descriptor]), s("regular-file"));
    assert_eq!(types("[method]descriptor.get-type", vec![root]), s("directory"));
}

// ── [method]descriptor.read-directory ─────────────────────────────

#[test]
fn read_directory_then_iterate_entries() {
    let dir = scratch_dir("read_directory");
    std::fs::write(dir.join("a.txt"), "").unwrap();
    std::fs::write(dir.join("b.txt"), "").unwrap();
    std::fs::create_dir(dir.join("sub")).unwrap();
    let root = open_test_root(&dir);

    let stream = types("[method]descriptor.read-directory", vec![root]);
    assert!(matches!(stream, Value::Object(_)), "read-directory returns a directory-entry-stream");

    let mut names = Vec::new();
    loop {
        let entry = types("[method]directory-entry-stream.read-directory-entry", vec![stream.clone()]);
        // option<directory-entry>: null = end-of-stream
        if matches!(entry, Value::Null) { break; }
        let name = match prop(&entry, "name") {
            Value::String(text) => text.to_string(),
            other => panic!("directory-entry.name expected string, got {:?}", other),
        };
        names.push(name);
        if names.len() > 50 { panic!("read-directory-entry didn't terminate"); }
    }
    names.sort();
    assert_eq!(names, vec!["a.txt", "b.txt", "sub"]);
}

// ── [method]descriptor.create-directory-at ────────────────────────

#[test]
fn create_directory_at_makes_subdir() {
    let dir = scratch_dir("create_dir");
    let root = open_test_root(&dir);
    let result = types("[method]descriptor.create-directory-at", vec![root, s("sub")]);
    assert!(is_error(&result).is_none(), "create-directory-at should succeed");
    assert!(dir.join("sub").is_dir(), "sub/ should exist on disk");
}

// ── [method]descriptor.unlink-file-at ─────────────────────────────

#[test]
fn unlink_file_at_removes_file() {
    let dir = scratch_dir("unlink_file");
    std::fs::write(dir.join("doomed"), "").unwrap();
    let root = open_test_root(&dir);
    let result = types("[method]descriptor.unlink-file-at", vec![root, s("doomed")]);
    assert!(is_error(&result).is_none());
    assert!(!dir.join("doomed").exists());
}

#[test]
fn unlink_file_at_on_directory_errors_is_directory() {
    let dir = scratch_dir("unlink_dir");
    std::fs::create_dir(dir.join("a-dir")).unwrap();
    let root = open_test_root(&dir);
    let result = types("[method]descriptor.unlink-file-at", vec![root, s("a-dir")]);
    assert_eq!(is_error(&result).as_deref(), Some("is-directory"));
}

// ── [method]descriptor.remove-directory-at ────────────────────────

#[test]
fn remove_directory_at_removes_empty_dir() {
    let dir = scratch_dir("rmdir_empty");
    std::fs::create_dir(dir.join("empty")).unwrap();
    let root = open_test_root(&dir);
    let result = types("[method]descriptor.remove-directory-at", vec![root, s("empty")]);
    assert!(is_error(&result).is_none());
    assert!(!dir.join("empty").exists());
}

#[test]
fn remove_directory_at_non_empty_errors_not_empty() {
    let dir = scratch_dir("rmdir_full");
    let sub = dir.join("full");
    std::fs::create_dir(&sub).unwrap();
    std::fs::write(sub.join("x"), "").unwrap();
    let root = open_test_root(&dir);
    let result = types("[method]descriptor.remove-directory-at", vec![root, s("full")]);
    assert_eq!(is_error(&result).as_deref(), Some("not-empty"));
}

// ── [method]descriptor.rename-at ──────────────────────────────────

#[test]
fn rename_at_moves_within_same_parent() {
    let dir = scratch_dir("rename_within");
    std::fs::write(dir.join("a"), "data").unwrap();
    let root = open_test_root(&dir);
    let result = types("[method]descriptor.rename-at", vec![
        root.clone(), s("a"), root, s("b"),
    ]);
    assert!(is_error(&result).is_none());
    assert!(!dir.join("a").exists());
    assert_eq!(std::fs::read_to_string(dir.join("b")).unwrap(), "data");
}

// ── [method]descriptor.is-same-object ─────────────────────────────

#[test]
fn is_same_object_true_for_same_descriptor() {
    let dir = scratch_dir("same_obj");
    let root = open_test_root(&dir);
    let same = types("[method]descriptor.is-same-object", vec![root.clone(), root]);
    assert_eq!(same, Value::Bool(true));
}

// ── [method]descriptor.read-via-stream + wasi:io/streams.read ─────

#[test]
fn read_via_stream_yields_input_stream() {
    let dir = scratch_dir("read_stream");
    let file = dir.join("payload.bin");
    std::fs::write(&file, b"hello, world").unwrap();
    let root = open_test_root(&dir);

    let descriptor = types("[method]descriptor.open-at", vec![
        root, Value::I32(0), s("payload.bin"), Value::I32(0), Value::I32(0),
    ]);
    // read-via-stream(offset=0)
    let stream = types("[method]descriptor.read-via-stream", vec![
        descriptor, Value::F64(0.0),
    ]);
    assert!(is_error(&stream).is_none(), "read-via-stream should succeed");
    assert!(matches!(stream, Value::Object(_)), "read-via-stream returns an input-stream");

    // Now read from it via wasi:io/streams.[method]input-stream.blocking-read
    let chunk = invoke(
        "wasi:io/streams",
        "[method]input-stream.blocking-read",
        vec![stream, Value::F64(64.0)],
    );
    if let Value::Object(object) = &chunk {
        let object = object.lock().unwrap();
        if let ObjectKind::Array(bytes) = &object.kind {
            let actual: Vec<u8> = bytes
                .iter()
                .filter_map(|value| match value {
                    Value::I32(n) => Some(*n as u8),
                    Value::F64(n) => Some(*n as u8),
                    _ => None,
                })
                .collect();
            assert_eq!(actual, b"hello, world");
            return;
        }
    }
    panic!("blocking-read should return a list<u8>, got {:?}", chunk);
}

#[allow(dead_code)]
fn _force_object_use(_: Object) {}
