use std::path::PathBuf;
use std::sync::atomic::{AtomicU64, Ordering};

use vybe_runtime::value::Value;
use vybe_runtime::{Chunk, Op, VM};
use vybe_runtime::capabilities::Capabilities;
use vybe_compiler::primitives::platforms::register_platforms;

fn scratch_dir(label: &str) -> PathBuf {
    static COUNTER: AtomicU64 = AtomicU64::new(0);
    let id = COUNTER.fetch_add(1, Ordering::SeqCst);
    let dir = std::env::temp_dir().join(format!(
        "vybe-wasi-fs-names-test-{}-{}-{}",
        std::process::id(),
        label,
        id
    ));
    let _ = std::fs::remove_dir_all(&dir);
    std::fs::create_dir_all(&dir).expect("scratch dir mkdir");
    dir
}

fn invoke(module: &str, name: &str, args: Vec<Value>) -> Value {
    let mut chunk = Chunk::new("<wasi-fs-names-test>");
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

fn directory_names(value: &Value) -> Vec<String> {
    let mut names = Vec::new();
    loop {
        let entry = types(
            "[method]directory-entry-stream.read-directory-entry",
            vec![value.clone()],
        );
        if matches!(entry, Value::Null) {
            break;
        }
        if let Value::String(name) = prop(&entry, "name") {
            names.push(name.to_string());
        }
        assert!(names.len() < 64, "directory stream did not terminate");
    }
    names.sort();
    names
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
    let stream = types("[method]descriptor.read-directory", vec![root]);
    assert_eq!(
        directory_names(&stream),
        vec![
            String::from(".hidden"),
            String::from("CamelCase.TXT"),
            String::from("multi.part.dir"),
            String::from("space dir"),
            String::from("space name.txt"),
        ]
    );
}
