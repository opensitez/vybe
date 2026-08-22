use std::path::PathBuf;
use std::sync::atomic::{AtomicU64, Ordering};

use vybe_compiler::primitives::platforms::register_platforms;
use vybe_runtime::capabilities::Capabilities;
use vybe_runtime::value::{ObjectKind, Value};
use vybe_runtime::{Chunk, Op, VM};

fn scratch_dir(label: &str) -> PathBuf {
    static COUNTER: AtomicU64 = AtomicU64::new(0);
    let id = COUNTER.fetch_add(1, Ordering::SeqCst);
    let dir = std::env::current_dir()
        .expect("cwd is the `.` preopen")
        .join("target/wasi-fs-tests")
        .join(format!(
        "vybe-wasi-fs-symlink-matrix-test-{}-{}-{}",
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
    let mut chunk = Chunk::new("<wasi-fs-symlink-matrix-test>");
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
                vm.set_global_owned(global_name.clone(), other);
                let ci = chunk.intern_string_constant(&global_name);
                chunk.emit_op_u16(Op::GLOBAL_GET, ci, 0);
            }
        }
    }
    chunk.emit_call(import_idx, argc, 0);
    chunk.emit_op(Op::RETURN, 0);

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

/// `open-flags::directory` — `proposals/WASI/proposals/filesystem/wit/types.wit`.
const OPEN_DIRECTORY: i32 = 2;

/// Element 0 of the first `tuple<descriptor, string>` `get-directories` answers.
fn first_preopen_descriptor(preopens: &Value) -> Value {
    let Value::Object(list) = preopens else {
        panic!("get-directories must answer a list");
    };
    let list = list.lock().unwrap();
    let ObjectKind::Array(entries) = &list.kind else {
        panic!("get-directories must answer list<tuple<descriptor, string>>");
    };
    let Some(Value::Object(pair)) = entries.first() else {
        panic!("no preopened directory — the guest has no capability at all");
    };
    let pair = pair.lock().unwrap();
    let ObjectKind::Array(pair) = &pair.kind else {
        panic!("each preopen is a tuple<descriptor, string>");
    };
    pair.first().cloned().expect("descriptor is element 0")
}

/// Open a descriptor for the scratch directory, through the REAL capability
/// path: `preopens.get-directories` then `open-at`.
///
/// This used to call `__test_open_root`, a host function that minted a
/// descriptor for any absolute path. It was registered inside
/// `wasi:filesystem/types` — so importable by any guest — and the WIT declares
/// no such function. Scratch directories therefore moved under the `.` preopen
/// (the CWD), which is the only place `open-at` can reach.
fn open_test_root(dir: &PathBuf) -> Value {
    let cwd = std::env::current_dir().expect("cwd is the `.` preopen");
    let rel = dir
        .strip_prefix(&cwd)
        .expect("scratch dirs live under the `.` preopen so `open-at` can reach them");

    let preopens = invoke("wasi:filesystem/preopens", "get-directories", vec![]);
    let root = first_preopen_descriptor(&preopens);

    let opened = invoke(
        "wasi:filesystem/types",
        "[method]descriptor.open-at",
        vec![
            root,
            Value::I32(0),
            s(rel.to_str().expect("scratch path is utf-8")),
            Value::I32(OPEN_DIRECTORY),
            Value::I32(0),
        ],
    );
    assert!(
        is_error(&opened).is_none(),
        "opening the scratch dir from the preopen failed: {:?}",
        is_error(&opened)
    );
    opened
}

/// `read-directory: func() -> tuple<stream<directory-entry>,
///                                  future<result<_, error-code>>>`.
///
/// The SHAPE half only — that the answer is the declared tuple rather than the
/// 0.2 `directory-entry-stream` resource. The ENTRIES are asserted by each
/// caller through `stream_drain::read_directory`, which drains the stream
/// inside the VM that opened it; `invoke` cannot, because a stream end is an
/// index into a handle table that dies with the call.
fn assert_read_directory_tuple(result: &Value) {
    assert!(
        is_error(result).is_none(),
        "read-directory failed: {:?}",
        is_error(result)
    );
    let Value::Object(parts) = result else {
        panic!("read-directory must answer a tuple, got {result:?}");
    };
    let parts = parts.lock().unwrap();
    let ObjectKind::Array(parts) = &parts.kind else {
        panic!("read-directory must answer tuple<stream<directory-entry>, future<...>>");
    };
    assert_eq!(
        parts.len(),
        2,
        "the entry stream AND the completion future — not a resource handle"
    );
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
    let bytes = crate::stream_drain::read_via_stream(&dir, "alias.txt", 0.0);
    assert_eq!(bytes, b"hello");
}

#[test]
fn read_directory_reports_file_symlink_entry_as_symbolic_link() {
    let dir = scratch_dir("dirent_file_link");
    std::fs::write(dir.join("target.txt"), b"hello").unwrap();
    std::os::unix::fs::symlink("target.txt", dir.join("alias.txt")).unwrap();
    let root = open_test_root(&dir);
    let result = types("[method]descriptor.read-directory", vec![root]);
    assert_read_directory_tuple(&result);

    // `symbolic-link`, NOT `regular-file`. `read-directory` reports each
    // entry's own type without following it — the link and its target are
    // different entries with different types, and a host that stat'd through
    // the link would report both as `regular-file` and lose the distinction
    // an enumerator exists to make.
    assert_eq!(
        crate::stream_drain::read_directory(&dir),
        vec![
            ("symbolic-link".to_string(), "alias.txt".to_string()),
            ("regular-file".to_string(), "target.txt".to_string()),
        ]
    );
}

#[test]
fn read_directory_reports_directory_symlink_entry_as_symbolic_link() {
    let dir = scratch_dir("dirent_dir_link");
    std::fs::create_dir_all(dir.join("nested")).unwrap();
    std::os::unix::fs::symlink("nested", dir.join("nested-link")).unwrap();
    let root = open_test_root(&dir);
    let result = types("[method]descriptor.read-directory", vec![root]);
    assert_read_directory_tuple(&result);

    // The same rule where the target is a DIRECTORY, which is the case a
    // recursive walker gets wrong: following `nested-link` because it reads as
    // `directory` is how a walker enters a cycle.
    assert_eq!(
        crate::stream_drain::read_directory(&dir),
        vec![
            ("directory".to_string(), "nested".to_string()),
            ("symbolic-link".to_string(), "nested-link".to_string()),
        ]
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
