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
use vybe_compiler::primitives::platforms::register_platforms;
use vybe_runtime::capabilities::Capabilities;
use vybe_runtime::value::{Object, ObjectKind, Value};
use vybe_runtime::{Chunk, Op, VM};

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

static TEST_GLOBAL_SEQ: std::sync::atomic::AtomicUsize = std::sync::atomic::AtomicUsize::new(0);

fn invoke(module: &str, name: &str, args: Vec<Value>) -> Value {
    let mut chunk = Chunk::new("<wasi-fs-test>");
    let import_idx = chunk.add_import(module, name);
    let argc = args.len() as u8;
    let mut vm = VM::new();
    register_platforms(&mut vm, &Capabilities::all());
    for value in args {
        match value {
            Value::I32(n) => chunk.emit_i32_const(n, 0),
            Value::I64(n) => chunk.emit_i64_const(n, 0),
            Value::F32(f) => chunk.emit_f32_const(f, 0),
            Value::F64(f) => chunk.emit_f64_const(f, 0),
            Value::Bool(b) => chunk.emit_bool_const(b, 0),
            Value::String(text) => chunk.emit_string_const(&text, 0),
            Value::Null => chunk.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, 0),
            other => {
                let global_name = format!(
                    "__test_arg_{}",
                    TEST_GLOBAL_SEQ.fetch_add(1, std::sync::atomic::Ordering::Relaxed)
                );
                vm.globals.insert(global_name.clone(), other);
                let ci = chunk.intern_string_constant(&global_name);
                chunk.emit_op_u16(Op::GLOBAL_GET, ci, 0);
            }
        }
    }
    chunk.emit_call(import_idx, argc, 0);
    chunk.emit_op(Op::RETURN, 0);

    vm.run(vec![chunk]).expect("VM run failed")
}

fn has_import(module: &str, name: &str) -> bool {
    let mut vm = VM::new();
    register_platforms(&mut vm, &Capabilities::all());
    vm.host_registry
        .contains_key(&(module.to_string(), name.to_string()))
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
    invoke(
        "wasi:filesystem/types",
        "__test_open_root",
        vec![s(dir.to_str().unwrap())],
    )
}

macro_rules! assert_wasi_error {
    ($value:expr, $code:expr) => {
        assert_eq!(is_error(&$value).as_deref(), Some($code));
    };
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
                        assert!(
                            matches!(pair_elements[0], Value::Object(_)),
                            "pair[0] descriptor"
                        );
                        assert!(
                            matches!(pair_elements[1], Value::String(_)),
                            "pair[1] path string"
                        );
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
    let result = types(
        "[method]descriptor.open-at",
        vec![
            root,
            Value::I32(0),
            s("hello.txt"),
            Value::I32(0),
            Value::I32(0),
        ],
    );
    assert!(
        is_error(&result).is_none(),
        "expected ok descriptor, got error {:?}",
        is_error(&result)
    );
    assert!(
        matches!(result, Value::Object(_)),
        "open-at returns a descriptor object"
    );
}

#[test]
fn open_at_missing_file_returns_no_entry_error() {
    let dir = scratch_dir("open_at_missing");
    let root = open_test_root(&dir);
    let result = types(
        "[method]descriptor.open-at",
        vec![
            root,
            Value::I32(0),
            s("nope.txt"),
            Value::I32(0),
            Value::I32(0),
        ],
    );
    assert_eq!(is_error(&result).as_deref(), Some("no-entry"));
}

#[test]
fn open_at_with_create_flag_creates_new_file() {
    let dir = scratch_dir("open_at_create");
    let root = open_test_root(&dir);
    // open-flags::create = bit 0 (per WIT: create, directory, exclusive, truncate)
    let result = types(
        "[method]descriptor.open-at",
        vec![
            root,
            Value::I32(0),
            s("new.txt"),
            Value::I32(1),
            Value::I32(0),
        ],
    );
    assert!(is_error(&result).is_none(), "open-at create should succeed");
    assert!(
        dir.join("new.txt").exists(),
        "file should be created on disk"
    );
}

#[test]
fn open_at_with_create_exclusive_on_existing_file_errors_exist() {
    let dir = scratch_dir("open_at_exclusive_exists");
    std::fs::write(dir.join("taken.txt"), "hi").unwrap();
    let root = open_test_root(&dir);

    let result = types(
        "[method]descriptor.open-at",
        vec![
            root,
            Value::I32(0),
            s("taken.txt"),
            Value::I32(1 | 4),
            Value::I32(0),
        ],
    );
    assert_eq!(is_error(&result).as_deref(), Some("exist"));
}

#[test]
fn open_at_with_directory_flag_on_file_errors_not_directory() {
    let dir = scratch_dir("open_at_directory_flag");
    std::fs::write(dir.join("plain.txt"), "hi").unwrap();
    let root = open_test_root(&dir);

    let result = types(
        "[method]descriptor.open-at",
        vec![
            root,
            Value::I32(0),
            s("plain.txt"),
            Value::I32(2),
            Value::I32(0),
        ],
    );
    assert_eq!(is_error(&result).as_deref(), Some("not-directory"));
}

#[test]
fn open_at_with_truncate_flag_clears_existing_file() {
    let dir = scratch_dir("open_at_truncate");
    std::fs::write(dir.join("truncate.txt"), "payload").unwrap();
    let root = open_test_root(&dir);

    let result = types(
        "[method]descriptor.open-at",
        vec![
            root,
            Value::I32(0),
            s("truncate.txt"),
            Value::I32(8),
            Value::I32(0),
        ],
    );
    assert!(is_error(&result).is_none());
    assert_eq!(
        std::fs::read_to_string(dir.join("truncate.txt")).unwrap(),
        ""
    );
}

#[test]
fn open_at_directory_flag_on_existing_directory_returns_directory_descriptor() {
    let dir = scratch_dir("open_at_directory");
    std::fs::create_dir(dir.join("subdir")).unwrap();
    let root = open_test_root(&dir);

    let descriptor = types(
        "[method]descriptor.open-at",
        vec![
            root,
            Value::I32(0),
            s("subdir"),
            Value::I32(2),
            Value::I32(0),
        ],
    );
    assert_eq!(
        types("[method]descriptor.get-type", vec![descriptor]),
        s("directory")
    );
}

#[test]
fn stat_at_returns_descriptor_stat() {
    let dir = scratch_dir("stat_at");
    let file = dir.join("sized.bin");
    std::fs::write(&file, b"abcdef").unwrap();
    let root = open_test_root(&dir);

    // stat-at(path-flags, path)
    let stats = types(
        "[method]descriptor.stat-at",
        vec![root, Value::I32(0), s("sized.bin")],
    );
    assert_eq!(prop(&stats, "size"), Value::F64(6.0));
    // descriptor-stat fields per WIT: type, link-count, size, mtime, atime, ctime
    assert_eq!(prop(&stats, "type"), s("regular-file"));
}

#[test]
fn stat_at_missing_file_returns_no_entry_error() {
    let dir = scratch_dir("stat_at_missing");
    let root = open_test_root(&dir);

    let stats = types(
        "[method]descriptor.stat-at",
        vec![root, Value::I32(0), s("missing.txt")],
    );
    assert_eq!(is_error(&stats).as_deref(), Some("no-entry"));
}

#[test]
fn stat_on_open_descriptor() {
    let dir = scratch_dir("stat_descriptor");
    let file = dir.join("d.bin");
    std::fs::write(&file, b"xx").unwrap();
    let root = open_test_root(&dir);

    let descriptor = types(
        "[method]descriptor.open-at",
        vec![
            root,
            Value::I32(0),
            s("d.bin"),
            Value::I32(0),
            Value::I32(0),
        ],
    );
    let stats = types("[method]descriptor.stat", vec![descriptor]);
    assert_eq!(prop(&stats, "size"), Value::F64(2.0));
    assert_eq!(prop(&stats, "type"), s("regular-file"));
}

#[test]
fn stat_on_deleted_descriptor_returns_no_entry() {
    let dir = scratch_dir("stat_deleted_descriptor");
    let file = dir.join("gone.bin");
    std::fs::write(&file, b"xx").unwrap();
    let root = open_test_root(&dir);

    let descriptor = types(
        "[method]descriptor.open-at",
        vec![
            root,
            Value::I32(0),
            s("gone.bin"),
            Value::I32(0),
            Value::I32(0),
        ],
    );
    std::fs::remove_file(file).unwrap();

    let stats = types("[method]descriptor.stat", vec![descriptor]);
    assert_eq!(is_error(&stats).as_deref(), Some("no-entry"));
}

#[test]
fn get_type_returns_descriptor_type() {
    let dir = scratch_dir("get_type");
    let file = dir.join("t.txt");
    std::fs::write(&file, "").unwrap();
    let root = open_test_root(&dir);

    let descriptor = types(
        "[method]descriptor.open-at",
        vec![
            root.clone(),
            Value::I32(0),
            s("t.txt"),
            Value::I32(0),
            Value::I32(0),
        ],
    );
    assert_eq!(
        types("[method]descriptor.get-type", vec![descriptor]),
        s("regular-file")
    );
    assert_eq!(
        types("[method]descriptor.get-type", vec![root]),
        s("directory")
    );
}

#[test]
fn get_type_after_file_deleted_returns_no_entry() {
    let dir = scratch_dir("get_type_deleted");
    let file = dir.join("gone.txt");
    std::fs::write(&file, "bye").unwrap();
    let root = open_test_root(&dir);

    let descriptor = types(
        "[method]descriptor.open-at",
        vec![
            root,
            Value::I32(0),
            s("gone.txt"),
            Value::I32(0),
            Value::I32(0),
        ],
    );
    std::fs::remove_file(file).unwrap();

    let ty = types("[method]descriptor.get-type", vec![descriptor]);
    assert_eq!(is_error(&ty).as_deref(), Some("no-entry"));
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
    assert!(
        matches!(stream, Value::Object(_)),
        "read-directory returns a directory-entry-stream"
    );

    let mut names = Vec::new();
    loop {
        let entry = types(
            "[method]directory-entry-stream.read-directory-entry",
            vec![stream.clone()],
        );
        // option<directory-entry>: null = end-of-stream
        if matches!(entry, Value::Null) {
            break;
        }
        let name = match prop(&entry, "name") {
            Value::String(text) => text.to_string(),
            other => panic!("directory-entry.name expected string, got {:?}", other),
        };
        names.push(name);
        if names.len() > 50 {
            panic!("read-directory-entry didn't terminate");
        }
    }
    names.sort();
    assert_eq!(names, vec!["a.txt", "b.txt", "sub"]);
}

#[test]
fn read_directory_on_file_descriptor_errors_not_directory() {
    let dir = scratch_dir("read_directory_file");
    std::fs::write(dir.join("plain.txt"), "hi").unwrap();
    let root = open_test_root(&dir);
    let file = types(
        "[method]descriptor.open-at",
        vec![
            root,
            Value::I32(0),
            s("plain.txt"),
            Value::I32(0),
            Value::I32(0),
        ],
    );

    let result = types("[method]descriptor.read-directory", vec![file]);
    assert_eq!(is_error(&result).as_deref(), Some("not-directory"));
}

#[test]
fn directory_entry_stream_returns_null_after_end_of_stream() {
    let dir = scratch_dir("read_directory_eof");
    let root = open_test_root(&dir);

    let stream = types("[method]descriptor.read-directory", vec![root]);
    assert!(matches!(
        types(
            "[method]directory-entry-stream.read-directory-entry",
            vec![stream.clone()]
        ),
        Value::Null
    ));
    assert!(matches!(
        types(
            "[method]directory-entry-stream.read-directory-entry",
            vec![stream]
        ),
        Value::Null
    ));
}

// ── [method]descriptor.create-directory-at ────────────────────────

#[test]
fn create_directory_at_makes_subdir() {
    let dir = scratch_dir("create_dir");
    let root = open_test_root(&dir);
    let result = types(
        "[method]descriptor.create-directory-at",
        vec![root, s("sub")],
    );
    assert!(
        is_error(&result).is_none(),
        "create-directory-at should succeed"
    );
    assert!(dir.join("sub").is_dir(), "sub/ should exist on disk");
}

#[test]
fn create_directory_at_existing_directory_errors_exist() {
    let dir = scratch_dir("create_dir_exists");
    std::fs::create_dir(dir.join("sub")).unwrap();
    let root = open_test_root(&dir);

    let result = types(
        "[method]descriptor.create-directory-at",
        vec![root, s("sub")],
    );
    assert_eq!(is_error(&result).as_deref(), Some("exist"));
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

#[test]
fn unlink_file_at_missing_path_errors_no_entry() {
    let dir = scratch_dir("unlink_missing");
    let root = open_test_root(&dir);
    let result = types(
        "[method]descriptor.unlink-file-at",
        vec![root, s("missing")],
    );
    assert_eq!(is_error(&result).as_deref(), Some("no-entry"));
}

// ── [method]descriptor.remove-directory-at ────────────────────────

#[test]
fn remove_directory_at_removes_empty_dir() {
    let dir = scratch_dir("rmdir_empty");
    std::fs::create_dir(dir.join("empty")).unwrap();
    let root = open_test_root(&dir);
    let result = types(
        "[method]descriptor.remove-directory-at",
        vec![root, s("empty")],
    );
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
    let result = types(
        "[method]descriptor.remove-directory-at",
        vec![root, s("full")],
    );
    assert_eq!(is_error(&result).as_deref(), Some("not-empty"));
}

#[test]
fn remove_directory_at_missing_path_errors_no_entry() {
    let dir = scratch_dir("rmdir_missing");
    let root = open_test_root(&dir);
    let result = types(
        "[method]descriptor.remove-directory-at",
        vec![root, s("ghost")],
    );
    assert_eq!(is_error(&result).as_deref(), Some("no-entry"));
}

// ── [method]descriptor.rename-at ──────────────────────────────────

#[test]
fn rename_at_moves_within_same_parent() {
    let dir = scratch_dir("rename_within");
    std::fs::write(dir.join("a"), "data").unwrap();
    let root = open_test_root(&dir);
    let result = types(
        "[method]descriptor.rename-at",
        vec![root.clone(), s("a"), root, s("b")],
    );
    assert!(is_error(&result).is_none());
    assert!(!dir.join("a").exists());
    assert_eq!(std::fs::read_to_string(dir.join("b")).unwrap(), "data");
}

#[test]
fn rename_at_across_directory_descriptors_moves_file() {
    let dir = scratch_dir("rename_across");
    let from = dir.join("from");
    let to = dir.join("to");
    std::fs::create_dir(&from).unwrap();
    std::fs::create_dir(&to).unwrap();
    std::fs::write(from.join("note.txt"), "payload").unwrap();
    let root = open_test_root(&dir);
    let from_dir = types(
        "[method]descriptor.open-at",
        vec![
            root.clone(),
            Value::I32(0),
            s("from"),
            Value::I32(2),
            Value::I32(0),
        ],
    );
    let to_dir = types(
        "[method]descriptor.open-at",
        vec![root, Value::I32(0), s("to"), Value::I32(2), Value::I32(0)],
    );

    let result = types(
        "[method]descriptor.rename-at",
        vec![from_dir, s("note.txt"), to_dir, s("moved.txt")],
    );
    assert!(is_error(&result).is_none());
    assert_eq!(
        std::fs::read_to_string(to.join("moved.txt")).unwrap(),
        "payload"
    );
}

#[test]
fn rename_at_missing_source_errors_no_entry() {
    let dir = scratch_dir("rename_missing");
    let root = open_test_root(&dir);

    let result = types(
        "[method]descriptor.rename-at",
        vec![root.clone(), s("missing.txt"), root, s("dest.txt")],
    );
    assert_eq!(is_error(&result).as_deref(), Some("no-entry"));
}

#[cfg(unix)]
#[test]
fn readlink_at_returns_symlink_target() {
    let dir = scratch_dir("readlink_target");
    std::fs::write(dir.join("target.txt"), "payload").unwrap();
    std::os::unix::fs::symlink("target.txt", dir.join("alias.txt")).unwrap();
    let root = open_test_root(&dir);

    let result = types("[method]descriptor.readlink-at", vec![root, s("alias.txt")]);
    assert_eq!(result, s("target.txt"));
}

#[test]
fn readlink_at_missing_path_errors_no_entry() {
    let dir = scratch_dir("readlink_missing");
    let root = open_test_root(&dir);

    let result = types(
        "[method]descriptor.readlink-at",
        vec![root, s("missing-link")],
    );
    assert_eq!(is_error(&result).as_deref(), Some("no-entry"));
}

// ── [method]descriptor.is-same-object ─────────────────────────────

#[test]
fn is_same_object_true_for_same_descriptor() {
    let dir = scratch_dir("same_obj");
    let root = open_test_root(&dir);
    let same = types(
        "[method]descriptor.is-same-object",
        vec![root.clone(), root],
    );
    assert_eq!(same, Value::Bool(true));
}

#[test]
fn is_same_object_true_for_distinct_descriptors_to_same_path() {
    let dir = scratch_dir("same_obj_path");
    std::fs::write(dir.join("same.txt"), "payload").unwrap();
    let root = open_test_root(&dir);

    let first = types(
        "[method]descriptor.open-at",
        vec![
            root.clone(),
            Value::I32(0),
            s("same.txt"),
            Value::I32(0),
            Value::I32(0),
        ],
    );
    let second = types(
        "[method]descriptor.open-at",
        vec![
            root,
            Value::I32(0),
            s("same.txt"),
            Value::I32(0),
            Value::I32(0),
        ],
    );

    let same = types("[method]descriptor.is-same-object", vec![first, second]);
    assert_eq!(same, Value::Bool(true));
}

#[test]
fn is_same_object_false_for_different_paths() {
    let dir = scratch_dir("same_obj_false");
    std::fs::write(dir.join("a.txt"), "a").unwrap();
    std::fs::write(dir.join("b.txt"), "b").unwrap();
    let root = open_test_root(&dir);

    let first = types(
        "[method]descriptor.open-at",
        vec![
            root.clone(),
            Value::I32(0),
            s("a.txt"),
            Value::I32(0),
            Value::I32(0),
        ],
    );
    let second = types(
        "[method]descriptor.open-at",
        vec![
            root,
            Value::I32(0),
            s("b.txt"),
            Value::I32(0),
            Value::I32(0),
        ],
    );

    let same = types("[method]descriptor.is-same-object", vec![first, second]);
    assert_eq!(same, Value::Bool(false));
}

// ── [method]descriptor.read-via-stream + wasi:io/streams.read ─────

#[test]
fn read_via_stream_yields_input_stream() {
    let dir = scratch_dir("read_stream");
    let file = dir.join("payload.bin");
    std::fs::write(&file, b"hello, world").unwrap();
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
    // read-via-stream(offset=0)
    let stream = types(
        "[method]descriptor.read-via-stream",
        vec![descriptor, Value::F64(0.0)],
    );
    assert!(
        is_error(&stream).is_none(),
        "read-via-stream should succeed"
    );
    assert!(
        matches!(stream, Value::Object(_)),
        "read-via-stream returns an input-stream"
    );

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

#[test]
fn input_stream_read_respects_offset_and_length() {
    let dir = scratch_dir("read_stream_offset");
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
    let stream = types(
        "[method]descriptor.read-via-stream",
        vec![descriptor, Value::F64(2.0)],
    );
    let chunk = invoke(
        "wasi:io/streams",
        "[method]input-stream.read",
        vec![stream, Value::F64(2.0)],
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
            assert_eq!(actual, b"cd");
            return;
        }
    }

    panic!(
        "input-stream.read should return a list<u8>, got {:?}",
        chunk
    );
}

#[test]
fn read_via_stream_on_directory_errors_is_directory() {
    let dir = scratch_dir("read_stream_directory");
    let root = open_test_root(&dir);

    let result = types(
        "[method]descriptor.read-via-stream",
        vec![root, Value::F64(0.0)],
    );
    assert_eq!(is_error(&result).as_deref(), Some("is-directory"));
}

#[test]
fn open_at_rejects_bad_descriptor() {
    let result = types(
        "[method]descriptor.open-at",
        vec![
            Value::Null,
            Value::I32(0),
            s("child.txt"),
            Value::I32(0),
            Value::I32(0),
        ],
    );
    assert_wasi_error!(result, "bad-descriptor");
}

#[test]
fn open_at_rejects_non_string_path_argument() {
    let dir = scratch_dir("open_at_invalid_path_arg");
    let root = open_test_root(&dir);
    let result = types(
        "[method]descriptor.open-at",
        vec![
            root,
            Value::I32(0),
            Value::I32(42),
            Value::I32(0),
            Value::I32(0),
        ],
    );
    assert_wasi_error!(result, "invalid");
}

#[test]
fn stat_rejects_bad_descriptor() {
    let result = types("[method]descriptor.stat", vec![Value::Null]);
    assert_wasi_error!(result, "bad-descriptor");
}

#[test]
fn stat_at_rejects_bad_descriptor() {
    let result = types(
        "[method]descriptor.stat-at",
        vec![Value::Null, Value::I32(0), s("x")],
    );
    assert_wasi_error!(result, "bad-descriptor");
}

#[test]
fn read_directory_rejects_bad_descriptor() {
    let result = types("[method]descriptor.read-directory", vec![Value::Null]);
    assert_wasi_error!(result, "bad-descriptor");
}

#[test]
fn directory_entry_stream_read_rejects_bad_descriptor() {
    let result = types(
        "[method]directory-entry-stream.read-directory-entry",
        vec![Value::Null],
    );
    assert_wasi_error!(result, "bad-descriptor");
}

#[test]
fn create_directory_at_rejects_bad_descriptor() {
    let result = types(
        "[method]descriptor.create-directory-at",
        vec![Value::Null, s("sub")],
    );
    assert_wasi_error!(result, "bad-descriptor");
}

#[test]
fn unlink_file_at_rejects_bad_descriptor() {
    let result = types(
        "[method]descriptor.unlink-file-at",
        vec![Value::Null, s("file")],
    );
    assert_wasi_error!(result, "bad-descriptor");
}

#[test]
fn remove_directory_at_rejects_bad_descriptor() {
    let result = types(
        "[method]descriptor.remove-directory-at",
        vec![Value::Null, s("dir")],
    );
    assert_wasi_error!(result, "bad-descriptor");
}

#[test]
fn rename_at_rejects_bad_source_descriptor() {
    let dir = scratch_dir("rename_bad_source");
    let root = open_test_root(&dir);
    let result = types(
        "[method]descriptor.rename-at",
        vec![Value::Null, s("a"), root, s("b")],
    );
    assert_wasi_error!(result, "bad-descriptor");
}

#[test]
fn rename_at_rejects_bad_target_descriptor() {
    let dir = scratch_dir("rename_bad_target");
    let root = open_test_root(&dir);
    let result = types(
        "[method]descriptor.rename-at",
        vec![root, s("a"), Value::Null, s("b")],
    );
    assert_wasi_error!(result, "bad-descriptor");
}

#[test]
fn readlink_at_rejects_bad_descriptor() {
    let result = types(
        "[method]descriptor.readlink-at",
        vec![Value::Null, s("link")],
    );
    assert_wasi_error!(result, "bad-descriptor");
}

#[test]
fn read_via_stream_rejects_bad_descriptor() {
    let result = types(
        "[method]descriptor.read-via-stream",
        vec![Value::Null, Value::F64(0.0)],
    );
    assert_wasi_error!(result, "bad-descriptor");
}

#[test]
fn input_stream_read_rejects_bad_descriptor() {
    let result = invoke(
        "wasi:io/streams",
        "[method]input-stream.read",
        vec![Value::Null, Value::F64(1.0)],
    );
    assert_wasi_error!(result, "bad-descriptor");
}

#[test]
fn input_stream_blocking_read_rejects_bad_descriptor() {
    let result = invoke(
        "wasi:io/streams",
        "[method]input-stream.blocking-read",
        vec![Value::Null, Value::F64(1.0)],
    );
    assert_wasi_error!(result, "bad-descriptor");
}

#[allow(dead_code)]
fn _force_object_use(_: Object) {}

#[test]
fn proposal_filesystem_preopens_surface_is_registered() {
    let expected = ["get-directories"];
    let missing = expected
        .into_iter()
        .filter(|name| !has_import("wasi:filesystem/preopens", name))
        .collect::<Vec<_>>();
    assert!(
        missing.is_empty(),
        "missing filesystem preopens imports: {missing:?}"
    );
}

#[test]
fn proposal_filesystem_descriptor_surface_is_registered() {
    let expected = [
        "[method]descriptor.read-via-stream",
        "[method]descriptor.write-via-stream",
        "[method]descriptor.append-via-stream",
        "[method]descriptor.advise",
        "[method]descriptor.sync-data",
        "[method]descriptor.get-flags",
        "[method]descriptor.get-type",
        "[method]descriptor.set-size",
        "[method]descriptor.set-times",
        "[method]descriptor.read-directory",
        "[method]descriptor.sync",
        "[method]descriptor.create-directory-at",
        "[method]descriptor.stat",
        "[method]descriptor.stat-at",
        "[method]descriptor.set-times-at",
        "[method]descriptor.link-at",
        "[method]descriptor.open-at",
        "[method]descriptor.readlink-at",
        "[method]descriptor.remove-directory-at",
        "[method]descriptor.rename-at",
        "[method]descriptor.symlink-at",
        "[method]descriptor.unlink-file-at",
        "[method]descriptor.is-same-object",
        "[method]descriptor.metadata-hash",
        "[method]descriptor.metadata-hash-at",
        "[method]directory-entry-stream.read-directory-entry",
    ];
    let missing = expected
        .into_iter()
        .filter(|name| !has_import("wasi:filesystem/types", name))
        .collect::<Vec<_>>();
    assert!(
        missing.is_empty(),
        "missing filesystem descriptor imports: {missing:?}"
    );
}
