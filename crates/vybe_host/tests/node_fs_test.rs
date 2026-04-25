//! Behaviour tests for `node:fs` host imports.
//!
//! Each test is the contract for a single Node fs API. The
//! implementation in `src/node/fs.rs` grows to satisfy these tests
//! (TDD); when a test fails, the diagnostic should describe the gap
//! between Node behaviour (the contract) and Vybe's current emit.
//!
//! Reference: <https://nodejs.org/api/fs.html>.

use std::path::PathBuf;
use std::sync::atomic::{AtomicU64, Ordering};
use vybe_bytecode::value::{Object, ObjectKind, Value};
use vybe_bytecode::{Chunk, Op, VM};
use vybe_host::{Capabilities, register_with_capabilities};

// ── Test scaffolding ──────────────────────────────────────────────

/// Fresh per-test scratch directory under the system temp root. Tests
/// can write/delete inside without colliding with each other or with
/// repo files.
fn scratch_dir(label: &str) -> PathBuf {
    static COUNTER: AtomicU64 = AtomicU64::new(0);
    let id = COUNTER.fetch_add(1, Ordering::SeqCst);
    let dir = std::env::temp_dir().join(format!(
        "vybe-node-fs-test-{}-{}-{}",
        std::process::id(),
        label,
        id
    ));
    let _ = std::fs::remove_dir_all(&dir);
    std::fs::create_dir_all(&dir).expect("scratch dir mkdir");
    dir
}

/// Invoke a `node:fs:<name>` host fn with `args` as positional
/// constants on the stack. Returns the host fn's return value.
fn call_fs(name: &str, args: Vec<Value>) -> Value {
    let mut chunk = Chunk::new("<node-fs-test>");
    let import_idx = chunk.add_import("node:fs", name);
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

/// Convenience: wrap a borrowed string as a `Value::String`.
fn s(text: &str) -> Value {
    Value::String(std::sync::Arc::from(text))
}

/// Pull the array-of-strings out of a return value.
fn array_strings(value: &Value) -> Vec<String> {
    match value {
        Value::Object(object) => {
            let object = object.lock().unwrap();
            if let ObjectKind::Array(elements) = &object.kind {
                return elements
                    .iter()
                    .map(|element| match element {
                        Value::String(text) => text.to_string(),
                        other => format!("{}", other),
                    })
                    .collect();
            }
            Vec::new()
        }
        _ => Vec::new(),
    }
}

/// Pull a property off an Ordinary object.
fn prop(value: &Value, key: &str) -> Value {
    if let Value::Object(object) = value {
        let object = object.lock().unwrap();
        if let Some(v) = object.properties.get(key) {
            return v.clone();
        }
    }
    Value::Null
}

/// Invoke a method on a returned object: pulls `obj.<method>`, then
/// invokes it with `obj` as the receiver. Mirrors what the VM does
/// for `obj.method()`.
fn invoke_method(receiver: &Value, method: &str) -> Value {
    let method_ref = prop(receiver, method);
    // Build a chunk that pushes the receiver + method + does call_ref 1.
    let mut chunk = Chunk::new("<node-fs-method-test>");
    let recv_const = chunk.add_constant(receiver.clone());
    let method_const = chunk.add_constant(method_ref);
    chunk.emit_op_u16(Op::CONST, method_const, 0);
    chunk.emit_op_u16(Op::CONST, recv_const, 0);
    chunk.emit_op_u8(Op::CALL_REF, 1, 0);
    chunk.emit_op(Op::RETURN, 0);

    let mut vm = VM::new();
    register_with_capabilities(&mut vm, &Capabilities::all());
    vm.run(vec![chunk]).expect("invoke_method VM run failed")
}

// ── readFileSync ──────────────────────────────────────────────────

#[test]
fn read_file_sync_returns_utf8_string() {
    let dir = scratch_dir("rfs_basic");
    let file = dir.join("hello.txt");
    std::fs::write(&file, "hello world").unwrap();

    let value = call_fs("readFileSync", vec![s(file.to_str().unwrap()), s("utf8")]);
    assert_eq!(value, s("hello world"));
}

#[test]
fn read_file_sync_default_encoding_returns_buffer_like_array() {
    // Node: with no encoding arg, returns a Buffer (byte array). Vybe
    // doesn't have a Buffer class yet; the contract here is that the
    // result is iterable byte-array-shaped. Test pinned to that.
    let dir = scratch_dir("rfs_default");
    let file = dir.join("bytes.bin");
    std::fs::write(&file, [0x68u8, 0x69]).unwrap();

    let value = call_fs("readFileSync", vec![s(file.to_str().unwrap())]);
    if let Value::Object(obj) = &value {
        let o = obj.lock().unwrap();
        if let ObjectKind::Array(elems) = &o.kind {
            assert_eq!(elems.len(), 2, "two bytes expected");
            return;
        }
    }
    panic!("readFileSync without encoding should return a byte array, got {:?}", value);
}

// ── writeFileSync ─────────────────────────────────────────────────

#[test]
fn write_file_sync_creates_new_file() {
    let dir = scratch_dir("wfs_create");
    let file = dir.join("out.txt");
    call_fs("writeFileSync", vec![s(file.to_str().unwrap()), s("written")]);
    assert_eq!(std::fs::read_to_string(&file).unwrap(), "written");
}

#[test]
fn write_file_sync_overwrites_existing() {
    let dir = scratch_dir("wfs_overwrite");
    let file = dir.join("out.txt");
    std::fs::write(&file, "old").unwrap();
    call_fs("writeFileSync", vec![s(file.to_str().unwrap()), s("new")]);
    assert_eq!(std::fs::read_to_string(&file).unwrap(), "new");
}

// ── appendFileSync ────────────────────────────────────────────────

#[test]
fn append_file_sync_appends_to_existing() {
    let dir = scratch_dir("afs_append");
    let file = dir.join("log.txt");
    std::fs::write(&file, "line1\n").unwrap();
    call_fs("appendFileSync", vec![s(file.to_str().unwrap()), s("line2\n")]);
    assert_eq!(std::fs::read_to_string(&file).unwrap(), "line1\nline2\n");
}

#[test]
fn append_file_sync_creates_when_missing() {
    let dir = scratch_dir("afs_create");
    let file = dir.join("log.txt");
    call_fs("appendFileSync", vec![s(file.to_str().unwrap()), s("first\n")]);
    assert_eq!(std::fs::read_to_string(&file).unwrap(), "first\n");
}

// ── existsSync ────────────────────────────────────────────────────

#[test]
fn exists_sync_true_for_existing_file() {
    let dir = scratch_dir("exists_file");
    let file = dir.join("here.txt");
    std::fs::write(&file, "").unwrap();
    let value = call_fs("existsSync", vec![s(file.to_str().unwrap())]);
    assert_eq!(value, Value::Bool(true));
}

#[test]
fn exists_sync_true_for_existing_directory() {
    let dir = scratch_dir("exists_dir");
    let value = call_fs("existsSync", vec![s(dir.to_str().unwrap())]);
    assert_eq!(value, Value::Bool(true));
}

#[test]
fn exists_sync_false_for_missing_path() {
    let dir = scratch_dir("exists_missing");
    let missing = dir.join("nope.txt");
    let value = call_fs("existsSync", vec![s(missing.to_str().unwrap())]);
    assert_eq!(value, Value::Bool(false));
}

// ── statSync ──────────────────────────────────────────────────────

#[test]
fn stat_sync_returns_size() {
    let dir = scratch_dir("stat_size");
    let file = dir.join("sized.bin");
    std::fs::write(&file, b"abcdef").unwrap();
    let stats = call_fs("statSync", vec![s(file.to_str().unwrap())]);
    assert_eq!(prop(&stats, "size"), Value::F64(6.0));
}

#[test]
fn stat_sync_is_file_method_returns_true_for_file() {
    let dir = scratch_dir("stat_isfile_true");
    let file = dir.join("a.txt");
    std::fs::write(&file, "").unwrap();
    let stats = call_fs("statSync", vec![s(file.to_str().unwrap())]);
    assert_eq!(invoke_method(&stats, "isFile"), Value::Bool(true));
    assert_eq!(invoke_method(&stats, "isDirectory"), Value::Bool(false));
}

#[test]
fn stat_sync_is_directory_method_returns_true_for_dir() {
    let dir = scratch_dir("stat_isdir");
    let stats = call_fs("statSync", vec![s(dir.to_str().unwrap())]);
    assert_eq!(invoke_method(&stats, "isDirectory"), Value::Bool(true));
    assert_eq!(invoke_method(&stats, "isFile"), Value::Bool(false));
}

#[test]
fn stat_sync_mtime_ms_is_finite_number() {
    let dir = scratch_dir("stat_mtime");
    let file = dir.join("t.txt");
    std::fs::write(&file, "").unwrap();
    let stats = call_fs("statSync", vec![s(file.to_str().unwrap())]);
    let mtime = prop(&stats, "mtimeMs");
    if let Value::F64(ms) = mtime {
        assert!(ms > 0.0 && ms.is_finite(), "mtimeMs should be a finite positive number, got {}", ms);
    } else {
        panic!("mtimeMs expected number, got {:?}", mtime);
    }
}

// ── readdirSync ───────────────────────────────────────────────────

#[test]
fn readdir_sync_returns_entry_names() {
    let dir = scratch_dir("readdir_basic");
    std::fs::write(dir.join("a.txt"), "").unwrap();
    std::fs::write(dir.join("b.txt"), "").unwrap();
    std::fs::create_dir(dir.join("sub")).unwrap();

    let value = call_fs("readdirSync", vec![s(dir.to_str().unwrap())]);
    let mut names = array_strings(&value);
    names.sort();
    assert_eq!(names, vec!["a.txt", "b.txt", "sub"]);
}

#[test]
fn readdir_sync_returns_empty_array_for_empty_dir() {
    let dir = scratch_dir("readdir_empty");
    let value = call_fs("readdirSync", vec![s(dir.to_str().unwrap())]);
    assert_eq!(array_strings(&value), Vec::<String>::new());
}

// ── mkdirSync ─────────────────────────────────────────────────────

#[test]
fn mkdir_sync_creates_single_level() {
    let dir = scratch_dir("mkdir_basic");
    let target = dir.join("new");
    call_fs("mkdirSync", vec![s(target.to_str().unwrap())]);
    assert!(target.is_dir(), "expected new/ to exist as dir");
}

#[test]
fn mkdir_sync_recursive_creates_intermediate() {
    let dir = scratch_dir("mkdir_recursive");
    let target = dir.join("a/b/c");
    let opts = {
        let mut o = Object::new();
        o.properties.insert("recursive".into(), Value::Bool(true));
        Value::Object(std::sync::Arc::new(std::sync::Mutex::new(o)))
    };
    call_fs("mkdirSync", vec![s(target.to_str().unwrap()), opts]);
    assert!(target.is_dir(), "expected a/b/c/ to exist as dir");
}

// ── unlinkSync ────────────────────────────────────────────────────

#[test]
fn unlink_sync_removes_file() {
    let dir = scratch_dir("unlink_basic");
    let file = dir.join("doomed.txt");
    std::fs::write(&file, "").unwrap();
    call_fs("unlinkSync", vec![s(file.to_str().unwrap())]);
    assert!(!file.exists(), "expected file removed");
}

// ── rmdirSync ─────────────────────────────────────────────────────

#[test]
fn rmdir_sync_removes_empty_directory() {
    let dir = scratch_dir("rmdir_basic");
    let target = dir.join("empty");
    std::fs::create_dir(&target).unwrap();
    call_fs("rmdirSync", vec![s(target.to_str().unwrap())]);
    assert!(!target.exists(), "expected empty dir removed");
}

// ── rmSync ────────────────────────────────────────────────────────

#[test]
fn rm_sync_removes_file() {
    let dir = scratch_dir("rm_file");
    let file = dir.join("doomed.txt");
    std::fs::write(&file, "").unwrap();
    call_fs("rmSync", vec![s(file.to_str().unwrap())]);
    assert!(!file.exists(), "expected file removed");
}

#[test]
fn rm_sync_recursive_removes_non_empty_directory() {
    let dir = scratch_dir("rm_recursive");
    let target = dir.join("tree");
    std::fs::create_dir(&target).unwrap();
    std::fs::write(target.join("a"), "").unwrap();
    std::fs::write(target.join("b"), "").unwrap();
    let opts = {
        let mut o = Object::new();
        o.properties.insert("recursive".into(), Value::Bool(true));
        Value::Object(std::sync::Arc::new(std::sync::Mutex::new(o)))
    };
    call_fs("rmSync", vec![s(target.to_str().unwrap()), opts]);
    assert!(!target.exists(), "expected tree/ removed");
}

// ── renameSync ────────────────────────────────────────────────────

#[test]
fn rename_sync_moves_file() {
    let dir = scratch_dir("rename_basic");
    let src = dir.join("a.txt");
    let dst = dir.join("b.txt");
    std::fs::write(&src, "data").unwrap();
    call_fs(
        "renameSync",
        vec![s(src.to_str().unwrap()), s(dst.to_str().unwrap())],
    );
    assert!(!src.exists(), "src should be gone");
    assert_eq!(std::fs::read_to_string(&dst).unwrap(), "data");
}

// ── copyFileSync ──────────────────────────────────────────────────

#[test]
fn copy_file_sync_copies_contents() {
    let dir = scratch_dir("copy_basic");
    let src = dir.join("src.txt");
    let dst = dir.join("dst.txt");
    std::fs::write(&src, "payload").unwrap();
    call_fs(
        "copyFileSync",
        vec![s(src.to_str().unwrap()), s(dst.to_str().unwrap())],
    );
    assert_eq!(std::fs::read_to_string(&dst).unwrap(), "payload");
    assert!(src.exists(), "src should still exist after copy");
}

// ── realpathSync ──────────────────────────────────────────────────

#[test]
fn realpath_sync_resolves_existing_path() {
    let dir = scratch_dir("realpath_basic");
    let file = dir.join("real.txt");
    std::fs::write(&file, "").unwrap();
    let value = call_fs("realpathSync", vec![s(file.to_str().unwrap())]);
    if let Value::String(text) = value {
        assert!(text.contains("real.txt"), "expected resolved path to contain real.txt, got {}", text);
    } else {
        panic!("realpathSync expected string, got {:?}", value);
    }
}

// ── accessSync ────────────────────────────────────────────────────

#[test]
fn access_sync_returns_undefined_for_existing_path() {
    // Node: returns undefined on success; throws on failure.
    let dir = scratch_dir("access_ok");
    let file = dir.join("a.txt");
    std::fs::write(&file, "").unwrap();
    let value = call_fs("accessSync", vec![s(file.to_str().unwrap())]);
    // Vybe represents JS undefined as Value::Undefined; null is also
    // acceptable since we don't strictly distinguish in many places.
    assert!(
        matches!(value, Value::Undefined | Value::Null),
        "accessSync of existing path should be undefined/null, got {:?}",
        value
    );
}

// ── readlinkSync / lstatSync (gated on Unix; skipped on Windows) ──

#[cfg(unix)]
#[test]
fn readlink_sync_returns_target() {
    let dir = scratch_dir("readlink_basic");
    let target = dir.join("real.txt");
    let link = dir.join("link.txt");
    std::fs::write(&target, "").unwrap();
    std::os::unix::fs::symlink(&target, &link).unwrap();

    let value = call_fs("readlinkSync", vec![s(link.to_str().unwrap())]);
    if let Value::String(text) = value {
        assert!(text.contains("real.txt"), "expected target path, got {}", text);
    } else {
        panic!("readlinkSync expected string, got {:?}", value);
    }
}

#[cfg(unix)]
#[test]
fn lstat_sync_reports_symbolic_link() {
    let dir = scratch_dir("lstat_symlink");
    let target = dir.join("real.txt");
    let link = dir.join("link.txt");
    std::fs::write(&target, "").unwrap();
    std::os::unix::fs::symlink(&target, &link).unwrap();

    let stats = call_fs("lstatSync", vec![s(link.to_str().unwrap())]);
    assert_eq!(invoke_method(&stats, "isSymbolicLink"), Value::Bool(true));
    assert_eq!(invoke_method(&stats, "isFile"), Value::Bool(false));
}

// ── truncateSync ──────────────────────────────────────────────────

#[test]
fn truncate_sync_default_truncates_to_zero() {
    let dir = scratch_dir("truncate_default");
    let file = dir.join("t.txt");
    std::fs::write(&file, "abcdef").unwrap();
    call_fs("truncateSync", vec![s(file.to_str().unwrap())]);
    assert_eq!(std::fs::metadata(&file).unwrap().len(), 0);
}

#[test]
fn truncate_sync_to_specified_length() {
    let dir = scratch_dir("truncate_n");
    let file = dir.join("t.txt");
    std::fs::write(&file, "abcdef").unwrap();
    call_fs(
        "truncateSync",
        vec![s(file.to_str().unwrap()), Value::F64(3.0)],
    );
    assert_eq!(std::fs::read_to_string(&file).unwrap(), "abc");
}
