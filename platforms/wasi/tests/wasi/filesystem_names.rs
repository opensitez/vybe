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
        "vybe-wasi-fs-names-test-{}-{}-{}",
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
    let mut chunk = Chunk::new("<wasi-fs-names-test>");
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

macro_rules! file_name_case {
    ($open_name:ident, $stat_name:ident, $unlink_name:ident, $label:expr, $name:expr) => {
        #[test]
        fn $open_name() {
            let dir = scratch_dir($label);
            std::fs::write(dir.join($name), "payload").unwrap();
            let root = open_test_root(&dir);
            let descriptor = types(
                "[method]descriptor.open-at",
                vec![root, Value::I32(0), s($name), Value::I32(0), Value::I32(0)],
            );
            assert!(is_error(&descriptor).is_none());
            assert_eq!(
                types("[method]descriptor.get-type", vec![descriptor]),
                s("regular-file")
            );
        }

        #[test]
        fn $stat_name() {
            let dir = scratch_dir(concat!($label, "_stat"));
            std::fs::write(dir.join($name), "payload").unwrap();
            let root = open_test_root(&dir);
            let stat = types(
                "[method]descriptor.stat-at",
                vec![root, Value::I32(0), s($name)],
            );
            assert_eq!(prop(&stat, "size"), Value::F64(7.0));
            assert_eq!(prop(&stat, "type"), s("regular-file"));
        }

        #[test]
        fn $unlink_name() {
            let dir = scratch_dir(concat!($label, "_unlink"));
            std::fs::write(dir.join($name), "payload").unwrap();
            let root = open_test_root(&dir);
            let result = types("[method]descriptor.unlink-file-at", vec![root, s($name)]);
            assert!(is_error(&result).is_none());
            assert!(!dir.join($name).exists());
        }
    };
}

macro_rules! dir_name_case {
    ($create_name:ident, $open_name:ident, $rename_name:ident, $label:expr, $name:expr, $renamed:expr) => {
        #[test]
        fn $create_name() {
            let dir = scratch_dir($label);
            let root = open_test_root(&dir);
            let result = types(
                "[method]descriptor.create-directory-at",
                vec![root, s($name)],
            );
            assert!(is_error(&result).is_none());
            assert!(dir.join($name).is_dir());
        }

        #[test]
        fn $open_name() {
            let dir = scratch_dir(concat!($label, "_open"));
            std::fs::create_dir(dir.join($name)).unwrap();
            let root = open_test_root(&dir);
            let descriptor = types(
                "[method]descriptor.open-at",
                vec![root, Value::I32(0), s($name), Value::I32(2), Value::I32(0)],
            );
            assert_eq!(
                types("[method]descriptor.get-type", vec![descriptor]),
                s("directory")
            );
        }

        #[test]
        fn $rename_name() {
            let dir = scratch_dir(concat!($label, "_rename"));
            std::fs::create_dir(dir.join($name)).unwrap();
            let root = open_test_root(&dir);
            let result = types(
                "[method]descriptor.rename-at",
                vec![root.clone(), s($name), root, s($renamed)],
            );
            assert!(is_error(&result).is_none());
            assert!(dir.join($renamed).is_dir());
            assert!(!dir.join($name).exists());
        }
    };
}

file_name_case!(
    open_hidden_file_name,
    stat_hidden_file_name,
    unlink_hidden_file_name,
    "hidden_file",
    ".hidden"
);
file_name_case!(
    open_spaced_file_name,
    stat_spaced_file_name,
    unlink_spaced_file_name,
    "spaced_file",
    "space name.txt"
);
file_name_case!(
    open_mixed_case_file_name,
    stat_mixed_case_file_name,
    unlink_mixed_case_file_name,
    "mixed_case_file",
    "CamelCase.TXT"
);
file_name_case!(
    open_multi_dot_file_name,
    stat_multi_dot_file_name,
    unlink_multi_dot_file_name,
    "multi_dot_file",
    "archive.tar.gz"
);
file_name_case!(
    open_numeric_file_name,
    stat_numeric_file_name,
    unlink_numeric_file_name,
    "numeric_file",
    "12345.bin"
);

dir_name_case!(
    create_hidden_directory_name,
    open_hidden_directory_name,
    rename_hidden_directory_name,
    "hidden_dir",
    ".config",
    ".config-renamed"
);
dir_name_case!(
    create_spaced_directory_name,
    open_spaced_directory_name,
    rename_spaced_directory_name,
    "spaced_dir",
    "space dir",
    "space dir renamed"
);
dir_name_case!(
    create_mixed_case_directory_name,
    open_mixed_case_directory_name,
    rename_mixed_case_directory_name,
    "mixed_case_dir",
    "MixedCaseDir",
    "MixedCaseDirRenamed"
);
dir_name_case!(
    create_multi_dot_directory_name,
    open_multi_dot_directory_name,
    rename_multi_dot_directory_name,
    "multi_dot_dir",
    "multi.part.dir",
    "multi.part.dir.renamed"
);
dir_name_case!(
    create_numeric_directory_name,
    open_numeric_directory_name,
    rename_numeric_directory_name,
    "numeric_dir",
    "12345dir",
    "12345dir-renamed"
);

#[test]
fn read_directory_lists_special_file_and_directory_names() {
    let dir = scratch_dir("read_special_names");
    std::fs::write(dir.join(".hidden"), "x").unwrap();
    std::fs::write(dir.join("space name.txt"), "x").unwrap();
    std::fs::write(dir.join("CamelCase.TXT"), "x").unwrap();
    std::fs::create_dir(dir.join("space dir")).unwrap();
    std::fs::create_dir(dir.join("multi.part.dir")).unwrap();
    let root = open_test_root(&dir);
    let result = types("[method]descriptor.read-directory", vec![root]);
    assert_read_directory_tuple(&result);

    // Every name here is one a naive lift mangles differently: a leading dot
    // is not a hidden-file flag to be filtered, an embedded space is not a
    // separator, and case is preserved rather than folded. They arrive as a
    // (ptr, length) pair decoded from linear memory, so a length read at the
    // wrong offset truncates or over-reads exactly one of them.
    assert_eq!(
        crate::stream_drain::read_directory_names(&dir),
        vec![
            ".hidden".to_string(),
            "CamelCase.TXT".to_string(),
            "multi.part.dir".to_string(),
            "space dir".to_string(),
            "space name.txt".to_string(),
        ]
    );
}

/// A non-ASCII entry name survives the trip through linear memory.
///
/// The `name` field is a UTF-8 (ptr, length) pair and `length` is in BYTES,
/// not characters. Every name in the case above is ASCII, where the two
/// coincide — so a lift that treated the length as a character count would
/// pass all five and truncate the moment a real user named a file "café".
#[test]
fn read_directory_preserves_multibyte_entry_names() {
    let dir = scratch_dir("read_multibyte_names");
    std::fs::write(dir.join("café.txt"), "x").unwrap();
    std::fs::create_dir(dir.join("日本語")).unwrap();

    assert_eq!(
        crate::stream_drain::read_directory(&dir),
        vec![
            ("regular-file".to_string(), "café.txt".to_string()),
            ("directory".to_string(), "日本語".to_string()),
        ]
    );
}
