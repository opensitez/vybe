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
        "vybe-wasi-fs-paths-test-{}-{}-{}",
        std::process::id(),
        label,
        id
    ));
    let _ = std::fs::remove_dir_all(&dir);
    std::fs::create_dir_all(&dir).expect("scratch dir mkdir");
    dir
}

fn invoke(module: &str, name: &str, args: Vec<Value>) -> Value {
    let mut chunk = Chunk::new("<wasi-fs-paths-test>");
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

fn open_directory(parent: Value, child: &str) -> Value {
    let descriptor = types(
        "[method]descriptor.open-at",
        vec![
            parent,
            Value::I32(0),
            s(child),
            Value::I32(2),
            Value::I32(0),
        ],
    );
    assert!(
        is_error(&descriptor).is_none(),
        "directory open should succeed"
    );
    descriptor
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
        assert!(entries.len() < 64, "directory stream did not terminate");
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
fn open_at_root_can_open_nested_file_path() {
    let dir = scratch_dir("open_nested_root");
    std::fs::create_dir_all(dir.join("nested/deeper")).unwrap();
    std::fs::write(dir.join("nested/deeper/file.txt"), "payload").unwrap();
    let root = open_test_root(&dir);

    let descriptor = types(
        "[method]descriptor.open-at",
        vec![
            root,
            Value::I32(0),
            s("nested/deeper/file.txt"),
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
fn open_at_directory_descriptor_can_open_child_relative_path() {
    let dir = scratch_dir("open_nested_relative");
    std::fs::create_dir_all(dir.join("nested/deeper")).unwrap();
    std::fs::write(dir.join("nested/deeper/file.txt"), "payload").unwrap();
    let root = open_test_root(&dir);
    let nested = open_directory(root, "nested");

    let descriptor = types(
        "[method]descriptor.open-at",
        vec![
            nested,
            Value::I32(0),
            s("deeper/file.txt"),
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
fn open_at_create_creates_file_beneath_open_subdirectory() {
    let dir = scratch_dir("create_in_subdir");
    std::fs::create_dir_all(dir.join("nested")).unwrap();
    let root = open_test_root(&dir);
    let nested = open_directory(root, "nested");

    let descriptor = types(
        "[method]descriptor.open-at",
        vec![
            nested,
            Value::I32(0),
            s("created.txt"),
            Value::I32(1),
            Value::I32(0),
        ],
    );
    assert!(is_error(&descriptor).is_none());
    assert!(dir.join("nested/created.txt").exists());
}

#[test]
fn open_at_create_fails_when_intermediate_parent_missing() {
    let dir = scratch_dir("create_missing_parent");
    let root = open_test_root(&dir);
    let result = types(
        "[method]descriptor.open-at",
        vec![
            root,
            Value::I32(0),
            s("missing/child.txt"),
            Value::I32(1),
            Value::I32(0),
        ],
    );
    assert_wasi_error!(result, "no-entry");
}

#[test]
fn open_at_truncate_only_affects_target_nested_file() {
    let dir = scratch_dir("truncate_nested_target");
    std::fs::create_dir_all(dir.join("nested")).unwrap();
    std::fs::write(dir.join("nested/a.txt"), "aaaa").unwrap();
    std::fs::write(dir.join("nested/b.txt"), "bbbb").unwrap();
    let root = open_test_root(&dir);

    let result = types(
        "[method]descriptor.open-at",
        vec![
            root,
            Value::I32(0),
            s("nested/a.txt"),
            Value::I32(8),
            Value::I32(0),
        ],
    );
    assert!(is_error(&result).is_none());
    assert_eq!(
        std::fs::read_to_string(dir.join("nested/a.txt")).unwrap(),
        ""
    );
    assert_eq!(
        std::fs::read_to_string(dir.join("nested/b.txt")).unwrap(),
        "bbbb"
    );
}

#[test]
fn open_at_on_nested_directory_returns_directory_descriptor() {
    let dir = scratch_dir("open_nested_directory");
    std::fs::create_dir_all(dir.join("nested/deeper")).unwrap();
    let root = open_test_root(&dir);

    let descriptor = types(
        "[method]descriptor.open-at",
        vec![
            root,
            Value::I32(0),
            s("nested/deeper"),
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
fn stat_at_reports_nested_file_size_from_root_descriptor() {
    let dir = scratch_dir("stat_nested_root");
    std::fs::create_dir_all(dir.join("nested")).unwrap();
    std::fs::write(dir.join("nested/file.bin"), b"abcdef").unwrap();
    let root = open_test_root(&dir);

    let stat = types(
        "[method]descriptor.stat-at",
        vec![root, Value::I32(0), s("nested/file.bin")],
    );
    assert_eq!(prop(&stat, "size"), Value::F64(6.0));
    assert_eq!(prop(&stat, "type"), s("regular-file"));
}

#[test]
fn stat_at_reports_nested_file_size_from_subdirectory_descriptor() {
    let dir = scratch_dir("stat_nested_subdir");
    std::fs::create_dir_all(dir.join("nested")).unwrap();
    std::fs::write(dir.join("nested/file.bin"), b"abcdefghi").unwrap();
    let root = open_test_root(&dir);
    let nested = open_directory(root, "nested");

    let stat = types(
        "[method]descriptor.stat-at",
        vec![nested, Value::I32(0), s("file.bin")],
    );
    assert_eq!(prop(&stat, "size"), Value::F64(9.0));
    assert_eq!(prop(&stat, "type"), s("regular-file"));
}

#[test]
fn stat_at_reports_nested_directory_type() {
    let dir = scratch_dir("stat_nested_directory");
    std::fs::create_dir_all(dir.join("nested/deeper")).unwrap();
    let root = open_test_root(&dir);

    let stat = types(
        "[method]descriptor.stat-at",
        vec![root, Value::I32(0), s("nested/deeper")],
    );
    assert_eq!(prop(&stat, "type"), s("directory"));
}

#[test]
fn stat_at_rejects_non_string_path_argument() {
    let dir = scratch_dir("stat_invalid_path_arg");
    let root = open_test_root(&dir);
    let result = types(
        "[method]descriptor.stat-at",
        vec![root, Value::I32(0), Value::I32(7)],
    );
    assert_wasi_error!(result, "invalid");
}

#[test]
fn read_directory_on_nested_subdirectory_lists_only_direct_children() {
    let dir = scratch_dir("read_nested_subdir");
    std::fs::create_dir_all(dir.join("nested/deeper")).unwrap();
    std::fs::write(dir.join("nested/file.txt"), "payload").unwrap();
    std::fs::write(dir.join("nested/deeper/leaf.txt"), "payload").unwrap();
    let root = open_test_root(&dir);
    let nested = open_directory(root, "nested");

    let stream = types("[method]descriptor.read-directory", vec![nested]);
    assert_eq!(
        directory_entries(stream),
        vec![
            (String::from("deeper"), String::from("directory")),
            (String::from("file.txt"), String::from("regular-file")),
        ]
    );
}

#[test]
fn read_directory_reports_file_and_directory_entry_types() {
    let dir = scratch_dir("entry_types");
    std::fs::create_dir_all(dir.join("nested/dir-child")).unwrap();
    std::fs::write(dir.join("nested/file-child.txt"), "payload").unwrap();
    let root = open_test_root(&dir);
    let nested = open_directory(root, "nested");

    let stream = types("[method]descriptor.read-directory", vec![nested]);
    let entries = directory_entries(stream);
    assert!(entries.contains(&(String::from("dir-child"), String::from("directory"))));
    assert!(entries.contains(&(String::from("file-child.txt"), String::from("regular-file"))));
}

#[test]
fn create_directory_at_creates_child_in_open_subdirectory() {
    let dir = scratch_dir("mkdir_in_subdir");
    std::fs::create_dir_all(dir.join("nested")).unwrap();
    let root = open_test_root(&dir);
    let nested = open_directory(root, "nested");

    let result = types(
        "[method]descriptor.create-directory-at",
        vec![nested, s("new-child")],
    );
    assert!(is_error(&result).is_none());
    assert!(dir.join("nested/new-child").is_dir());
}

#[test]
fn create_directory_at_fails_when_parent_component_is_missing() {
    let dir = scratch_dir("mkdir_missing_parent");
    let root = open_test_root(&dir);
    let result = types(
        "[method]descriptor.create-directory-at",
        vec![root, s("missing/child")],
    );
    assert_wasi_error!(result, "no-entry");
}

#[test]
fn create_directory_at_existing_file_returns_exist() {
    let dir = scratch_dir("mkdir_existing_file");
    std::fs::write(dir.join("taken"), "payload").unwrap();
    let root = open_test_root(&dir);
    let result = types(
        "[method]descriptor.create-directory-at",
        vec![root, s("taken")],
    );
    assert_wasi_error!(result, "exist");
}

#[test]
fn create_directory_at_rejects_non_string_path_argument() {
    let dir = scratch_dir("mkdir_invalid_arg");
    let root = open_test_root(&dir);
    let result = types(
        "[method]descriptor.create-directory-at",
        vec![root, Value::I32(9)],
    );
    assert_wasi_error!(result, "invalid");
}

#[test]
fn unlink_file_at_removes_nested_file_from_open_subdirectory() {
    let dir = scratch_dir("unlink_nested_file");
    std::fs::create_dir_all(dir.join("nested")).unwrap();
    std::fs::write(dir.join("nested/doomed.txt"), "payload").unwrap();
    let root = open_test_root(&dir);
    let nested = open_directory(root, "nested");

    let result = types(
        "[method]descriptor.unlink-file-at",
        vec![nested, s("doomed.txt")],
    );
    assert!(is_error(&result).is_none());
    assert!(!dir.join("nested/doomed.txt").exists());
}

#[test]
fn unlink_file_at_missing_parent_component_returns_no_entry() {
    let dir = scratch_dir("unlink_missing_parent");
    let root = open_test_root(&dir);
    let result = types(
        "[method]descriptor.unlink-file-at",
        vec![root, s("missing/file.txt")],
    );
    assert_wasi_error!(result, "no-entry");
}

#[test]
fn unlink_file_at_rejects_non_string_path_argument() {
    let dir = scratch_dir("unlink_invalid_arg");
    let root = open_test_root(&dir);
    let result = types(
        "[method]descriptor.unlink-file-at",
        vec![root, Value::I32(4)],
    );
    assert_wasi_error!(result, "invalid");
}

#[test]
fn remove_directory_at_removes_nested_empty_directory_from_open_subdirectory() {
    let dir = scratch_dir("rmdir_nested_empty");
    std::fs::create_dir_all(dir.join("nested/empty-child")).unwrap();
    let root = open_test_root(&dir);
    let nested = open_directory(root, "nested");

    let result = types(
        "[method]descriptor.remove-directory-at",
        vec![nested, s("empty-child")],
    );
    assert!(is_error(&result).is_none());
    assert!(!dir.join("nested/empty-child").exists());
}

#[test]
fn remove_directory_at_file_path_returns_not_directory() {
    let dir = scratch_dir("rmdir_file_path");
    std::fs::write(dir.join("plain.txt"), "payload").unwrap();
    let root = open_test_root(&dir);
    let result = types(
        "[method]descriptor.remove-directory-at",
        vec![root, s("plain.txt")],
    );
    assert_wasi_error!(result, "not-directory");
}

#[test]
fn remove_directory_at_rejects_non_string_path_argument() {
    let dir = scratch_dir("rmdir_invalid_arg");
    let root = open_test_root(&dir);
    let result = types(
        "[method]descriptor.remove-directory-at",
        vec![root, Value::I32(11)],
    );
    assert_wasi_error!(result, "invalid");
}

#[test]
fn rename_at_moves_nested_file_between_open_subdirectories() {
    let dir = scratch_dir("rename_between_subdirs");
    std::fs::create_dir_all(dir.join("from")).unwrap();
    std::fs::create_dir_all(dir.join("to")).unwrap();
    std::fs::write(dir.join("from/note.txt"), "payload").unwrap();
    let root = open_test_root(&dir);
    let from = open_directory(root.clone(), "from");
    let to = open_directory(root, "to");

    let result = types(
        "[method]descriptor.rename-at",
        vec![from, s("note.txt"), to, s("moved.txt")],
    );
    assert!(is_error(&result).is_none());
    assert_eq!(
        std::fs::read_to_string(dir.join("to/moved.txt")).unwrap(),
        "payload"
    );
}

#[test]
fn rename_at_can_rename_directory_entry_within_same_parent() {
    let dir = scratch_dir("rename_directory_entry");
    std::fs::create_dir_all(dir.join("old-name")).unwrap();
    let root = open_test_root(&dir);

    let result = types(
        "[method]descriptor.rename-at",
        vec![root.clone(), s("old-name"), root, s("new-name")],
    );
    assert!(is_error(&result).is_none());
    assert!(dir.join("new-name").is_dir());
    assert!(!dir.join("old-name").exists());
}

#[test]
fn rename_at_rejects_non_string_source_path_argument() {
    let dir = scratch_dir("rename_invalid_src");
    let root = open_test_root(&dir);
    let result = types(
        "[method]descriptor.rename-at",
        vec![root.clone(), Value::I32(1), root, s("dest")],
    );
    assert_wasi_error!(result, "invalid");
}

#[test]
fn rename_at_rejects_non_string_target_path_argument() {
    let dir = scratch_dir("rename_invalid_dst");
    let root = open_test_root(&dir);
    let result = types(
        "[method]descriptor.rename-at",
        vec![root.clone(), s("src"), root, Value::I32(2)],
    );
    assert_wasi_error!(result, "invalid");
}

#[test]
fn get_type_on_nested_directory_descriptor_returns_directory() {
    let dir = scratch_dir("get_type_nested_dir");
    std::fs::create_dir_all(dir.join("nested/inner")).unwrap();
    let root = open_test_root(&dir);
    let nested = open_directory(root, "nested/inner");
    assert_eq!(
        types("[method]descriptor.get-type", vec![nested]),
        s("directory")
    );
}

#[test]
fn is_same_object_is_false_for_root_and_child_descriptor() {
    let dir = scratch_dir("same_object_root_child");
    std::fs::write(dir.join("file.txt"), "payload").unwrap();
    let root = open_test_root(&dir);
    let child = types(
        "[method]descriptor.open-at",
        vec![
            root.clone(),
            Value::I32(0),
            s("file.txt"),
            Value::I32(0),
            Value::I32(0),
        ],
    );
    assert_eq!(
        types("[method]descriptor.is-same-object", vec![root, child]),
        Value::Bool(false)
    );
}

#[test]
fn read_via_stream_from_nested_file_observes_initial_offset() {
    let dir = scratch_dir("stream_nested_offset");
    std::fs::create_dir_all(dir.join("nested")).unwrap();
    std::fs::write(dir.join("nested/file.txt"), b"abcdef").unwrap();
    let root = open_test_root(&dir);
    let descriptor = types(
        "[method]descriptor.open-at",
        vec![
            root,
            Value::I32(0),
            s("nested/file.txt"),
            Value::I32(0),
            Value::I32(0),
        ],
    );

    let stream = types(
        "[method]descriptor.read-via-stream",
        vec![descriptor, Value::F64(3.0)],
    );
    let bytes = invoke(
        "wasi:io/streams",
        "[method]input-stream.read",
        vec![stream, Value::F64(3.0)],
    );
    assert_eq!(bytes_from_array(&bytes), b"def");
}

#[test]
fn input_stream_read_consumes_remaining_bytes_across_multiple_reads() {
    let dir = scratch_dir("stream_multiple_reads");
    std::fs::write(dir.join("payload.bin"), b"abcdefgh").unwrap();
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
        vec![descriptor, Value::F64(1.0)],
    );

    let first = invoke(
        "wasi:io/streams",
        "[method]input-stream.read",
        vec![stream.clone(), Value::F64(3.0)],
    );
    let second = invoke(
        "wasi:io/streams",
        "[method]input-stream.read",
        vec![stream, Value::F64(8.0)],
    );
    assert_eq!(bytes_from_array(&first), b"bcd");
    assert_eq!(bytes_from_array(&second), b"efgh");
}

#[test]
fn blocking_read_past_end_of_nested_file_returns_empty_array() {
    let dir = scratch_dir("stream_eof");
    std::fs::create_dir_all(dir.join("nested")).unwrap();
    std::fs::write(dir.join("nested/file.txt"), b"abc").unwrap();
    let root = open_test_root(&dir);
    let descriptor = types(
        "[method]descriptor.open-at",
        vec![
            root,
            Value::I32(0),
            s("nested/file.txt"),
            Value::I32(0),
            Value::I32(0),
        ],
    );
    let stream = types(
        "[method]descriptor.read-via-stream",
        vec![descriptor, Value::F64(3.0)],
    );

    let bytes = invoke(
        "wasi:io/streams",
        "[method]input-stream.blocking-read",
        vec![stream, Value::F64(5.0)],
    );
    assert_eq!(bytes_from_array(&bytes), Vec::<u8>::new());
}

#[cfg(unix)]
#[test]
fn readlink_at_reads_symlink_relative_target_from_open_subdirectory() {
    let dir = scratch_dir("readlink_subdir");
    std::fs::create_dir_all(dir.join("nested")).unwrap();
    std::fs::write(dir.join("nested/target.txt"), "payload").unwrap();
    std::os::unix::fs::symlink("target.txt", dir.join("nested/alias.txt")).unwrap();
    let root = open_test_root(&dir);
    let nested = open_directory(root, "nested");

    let result = types(
        "[method]descriptor.readlink-at",
        vec![nested, s("alias.txt")],
    );
    assert_eq!(result, s("target.txt"));
}

#[test]
fn readlink_at_rejects_non_string_path_argument() {
    let dir = scratch_dir("readlink_invalid_arg");
    let root = open_test_root(&dir);
    let result = types("[method]descriptor.readlink-at", vec![root, Value::I32(5)]);
    assert_wasi_error!(result, "invalid");
}
