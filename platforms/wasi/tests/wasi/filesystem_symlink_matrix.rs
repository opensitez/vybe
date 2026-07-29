use std::path::PathBuf;
use std::sync::atomic::{AtomicU64, Ordering};

use vybe_runtime::value::{ObjectKind, Value};
use vybe_runtime::{Chunk, Op, VM};
use vybe_runtime::capabilities::Capabilities;
use vybe_compiler::primitives::platforms::register_platforms;

fn scratch_dir(label: &str) -> PathBuf {
    static COUNTER: AtomicU64 = AtomicU64::new(0);
    let id = COUNTER.fetch_add(1, Ordering::SeqCst);
    let dir = std::env::temp_dir().join(format!(
        "vybe-wasi-fs-symlink-matrix-test-{}-{}-{}",
        std::process::id(),
        label,
        id
    ));
    let _ = std::fs::remove_dir_all(&dir);
    std::fs::create_dir_all(&dir).expect("scratch dir mkdir");
    dir
}

fn invoke(module: &str, name: &str, args: Vec<Value>) -> Value {
    let mut chunk = Chunk::new("<wasi-fs-symlink-matrix-test>");
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

fn io_streams(name: &str, args: Vec<Value>) -> Value {
    invoke("wasi:io/streams", name, args)
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

fn open_test_root(dir: &PathBuf) -> Value {
    invoke(
        "wasi:filesystem/types",
        "__test_open_root",
        vec![s(dir.to_str().unwrap())],
    )
}

fn bytes_from_array(value: &Value) -> Vec<u8> {
    let Value::Object(object) = value else {
        return Vec::new();
    };
    let object = object.lock().unwrap();
    let ObjectKind::Array(values) = &object.kind else {
        return Vec::new();
    };
    values
        .iter()
        .filter_map(|value| match value {
            Value::I32(byte) => Some(*byte as u8),
            Value::F64(byte) => Some(*byte as u8),
            _ => None,
        })
        .collect()
}

fn directory_entries(stream: Value) -> Vec<(String, String)> {
    let mut entries = Vec::new();
    loop {
        let entry = types(
            "[method]directory-entry-stream.read-directory-entry",
            vec![stream.clone()],
        );
        if matches!(entry, Value::Null) {
            break;
        }
        let name = match prop(&entry, "name") {
            Value::String(text) => text.to_string(),
            other => panic!("directory-entry.name expected string, got {:?}", other),
        };
        let kind = match prop(&entry, "type") {
            Value::String(text) => text.to_string(),
            other => panic!("directory-entry.type expected string, got {:?}", other),
        };
        entries.push((name, kind));
    }
    entries.sort();
    entries
}

macro_rules! assert_wasi_error {
    ($value:expr, $code:expr) => {
        assert_eq!(is_error(&$value).as_deref(), Some($code));
    };
}

#[test]
fn stat_at_symlink_to_file_reports_regular_file() {
    let dir = scratch_dir("stat_file_link");
    std::fs::write(dir.join("target.txt"), b"hello").unwrap();
    std::os::unix::fs::symlink("target.txt", dir.join("alias.txt")).unwrap();
    let root = open_test_root(&dir);

    let stat = types(
        "[method]descriptor.stat-at",
        vec![root, Value::I32(0), s("alias.txt")],
    );
    assert_eq!(prop(&stat, "type"), s("regular-file"));
    assert_eq!(prop(&stat, "size"), Value::F64(5.0));
}

#[test]
fn stat_at_symlink_to_directory_reports_directory() {
    let dir = scratch_dir("stat_dir_link");
    std::fs::create_dir_all(dir.join("nested/child")).unwrap();
    std::os::unix::fs::symlink("nested", dir.join("nested-link")).unwrap();
    let root = open_test_root(&dir);

    let stat = types(
        "[method]descriptor.stat-at",
        vec![root, Value::I32(0), s("nested-link")],
    );
    assert_eq!(prop(&stat, "type"), s("directory"));
}

#[test]
fn stat_at_symlink_follow_flag_matches_default_behavior() {
    let dir = scratch_dir("stat_follow_flag");
    std::fs::write(dir.join("target.txt"), b"hello").unwrap();
    std::os::unix::fs::symlink("target.txt", dir.join("alias.txt")).unwrap();
    let root = open_test_root(&dir);

    let default_stat = types(
        "[method]descriptor.stat-at",
        vec![root.clone(), Value::I32(0), s("alias.txt")],
    );
    let follow_flag_stat = types(
        "[method]descriptor.stat-at",
        vec![root, Value::I32(1), s("alias.txt")],
    );
    assert_eq!(prop(&default_stat, "type"), prop(&follow_flag_stat, "type"));
    assert_eq!(prop(&default_stat, "size"), prop(&follow_flag_stat, "size"));
}

#[test]
fn open_at_symlink_to_file_get_type_reports_regular_file() {
    let dir = scratch_dir("open_file_link");
    std::fs::write(dir.join("target.txt"), b"hello").unwrap();
    std::os::unix::fs::symlink("target.txt", dir.join("alias.txt")).unwrap();
    let root = open_test_root(&dir);

    let descriptor = types(
        "[method]descriptor.open-at",
        vec![
            root,
            Value::I32(0),
            s("alias.txt"),
            Value::I32(0),
            Value::I32(0),
        ],
    );
    assert!(is_error(&descriptor).is_none());
    assert_eq!(
        types("[method]descriptor.get-type", vec![descriptor]),
        s("regular-file")
    );
}

#[test]
fn open_at_symlink_to_directory_with_directory_flag_succeeds() {
    let dir = scratch_dir("open_dir_link");
    std::fs::create_dir_all(dir.join("nested")).unwrap();
    std::os::unix::fs::symlink("nested", dir.join("nested-link")).unwrap();
    let root = open_test_root(&dir);

    let descriptor = types(
        "[method]descriptor.open-at",
        vec![
            root,
            Value::I32(0),
            s("nested-link"),
            Value::I32(2),
            Value::I32(0),
        ],
    );
    assert!(is_error(&descriptor).is_none());
    assert_eq!(
        types("[method]descriptor.get-type", vec![descriptor]),
        s("directory")
    );
}

#[test]
fn read_via_stream_on_symlinked_file_reads_target_bytes() {
    let dir = scratch_dir("stream_file_link");
    std::fs::write(dir.join("target.txt"), b"hello").unwrap();
    std::os::unix::fs::symlink("target.txt", dir.join("alias.txt")).unwrap();
    let root = open_test_root(&dir);
    let descriptor = types(
        "[method]descriptor.open-at",
        vec![
            root,
            Value::I32(0),
            s("alias.txt"),
            Value::I32(0),
            Value::I32(0),
        ],
    );
    let stream = types(
        "[method]descriptor.read-via-stream",
        vec![descriptor, Value::F64(0.0)],
    );
    let bytes = io_streams(
        "[method]input-stream.blocking-read",
        vec![stream, Value::F64(5.0)],
    );
    assert_eq!(bytes_from_array(&bytes), b"hello");
}

#[test]
fn read_directory_reports_file_symlink_entry_as_symbolic_link() {
    let dir = scratch_dir("dirent_file_link");
    std::fs::write(dir.join("target.txt"), b"hello").unwrap();
    std::os::unix::fs::symlink("target.txt", dir.join("alias.txt")).unwrap();
    let root = open_test_root(&dir);
    let stream = types("[method]descriptor.read-directory", vec![root]);

    assert!(
        directory_entries(stream)
            .contains(&(String::from("alias.txt"), String::from("symbolic-link")))
    );
}

#[test]
fn read_directory_reports_directory_symlink_entry_as_symbolic_link() {
    let dir = scratch_dir("dirent_dir_link");
    std::fs::create_dir_all(dir.join("nested")).unwrap();
    std::os::unix::fs::symlink("nested", dir.join("nested-link")).unwrap();
    let root = open_test_root(&dir);
    let stream = types("[method]descriptor.read-directory", vec![root]);

    assert!(
        directory_entries(stream)
            .contains(&(String::from("nested-link"), String::from("symbolic-link")))
    );
}

#[test]
fn is_same_object_is_false_for_symlink_and_target_descriptors() {
    let dir = scratch_dir("same_object_link_vs_target");
    std::fs::write(dir.join("target.txt"), b"hello").unwrap();
    std::os::unix::fs::symlink("target.txt", dir.join("alias.txt")).unwrap();
    let root = open_test_root(&dir);

    let alias = types(
        "[method]descriptor.open-at",
        vec![
            root.clone(),
            Value::I32(0),
            s("alias.txt"),
            Value::I32(0),
            Value::I32(0),
        ],
    );
    let target = types(
        "[method]descriptor.open-at",
        vec![
            root,
            Value::I32(0),
            s("target.txt"),
            Value::I32(0),
            Value::I32(0),
        ],
    );

    assert_eq!(
        types("[method]descriptor.is-same-object", vec![alias, target]),
        Value::Bool(false)
    );
}

#[test]
fn open_at_broken_symlink_returns_no_entry() {
    let dir = scratch_dir("broken_link_open");
    std::os::unix::fs::symlink("missing.txt", dir.join("alias.txt")).unwrap();
    let root = open_test_root(&dir);

    let result = types(
        "[method]descriptor.open-at",
        vec![
            root,
            Value::I32(0),
            s("alias.txt"),
            Value::I32(0),
            Value::I32(0),
        ],
    );
    assert_wasi_error!(result, "no-entry");
}

#[test]
fn readlink_at_on_directory_symlink_returns_target_name() {
    let dir = scratch_dir("readlink_dir_link");
    std::fs::create_dir_all(dir.join("nested")).unwrap();
    std::os::unix::fs::symlink("nested", dir.join("nested-link")).unwrap();
    let root = open_test_root(&dir);

    let result = types(
        "[method]descriptor.readlink-at",
        vec![root, s("nested-link")],
    );
    assert_eq!(result, s("nested"));
}
