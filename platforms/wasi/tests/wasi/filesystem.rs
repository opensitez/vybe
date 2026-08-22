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

/// Where a scratch directory lives, RELATIVE to the `.` preopen.
///
/// It used to be `std::env::temp_dir()`, which no guest can reach: WASI grants
/// no ambient authority over absolute paths, so `/tmp/...` is only openable by
/// a host function that hands out descriptors for arbitrary paths — which is
/// exactly what `__test_open_root` was, and why it had to go. Under the CWD
/// preopen the same directory is reachable by `open-at`, the way a real guest
/// reaches anything.
///
/// `target/` because it is already ignored and `cargo clean` collects it.
const SCRATCH_ROOT: &str = "target/wasi-fs-tests";

/// The scratch directory's path relative to the preopen, and its absolute path.
///
/// Both are needed: the relative one is what `open-at` takes, the absolute one
/// is what `std::fs` uses to arrange the fixture.
fn scratch_rel(label: &str) -> String {
    static COUNTER: AtomicU64 = AtomicU64::new(0);
    let id = COUNTER.fetch_add(1, Ordering::SeqCst);
    format!(
        "{SCRATCH_ROOT}/vybe-wasi-fs-test-{}-{}-{}",
        std::process::id(),
        label,
        id
    )
}

fn scratch_dir(label: &str) -> PathBuf {
    let rel = scratch_rel(label);
    let dir = std::env::current_dir()
        .expect("cwd is the `.` preopen")
        .join(&rel);
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

/// Open a descriptor for the scratch directory — every test starts here.
///
/// This goes through the REAL capability path: `preopens.get-directories`
/// answers `list<tuple<descriptor, string>>`, and the scratch directory is
/// reached from that descriptor with `open-at`. It used to call a
/// `__test_open_root` host function that minted a descriptor for any absolute
/// path — a TEST FIXTURE registered inside `wasi:filesystem/types`, and so
/// callable by any guest that imported it. Its own comment argued no profile
/// routes imports to `__test_*`, but the namespace is the contract: a
/// conforming guest may import anything the interface declares, and the WIT
/// declares no such function.
///
/// Going through `open-at` is also strictly better coverage — these 38 tests
/// now exercise the preopen path that production code uses, instead of a
/// shortcut that existed only for them.
fn open_test_root(dir: &PathBuf) -> Value {
    let cwd = std::env::current_dir().expect("cwd is the `.` preopen");
    let rel = dir
        .strip_prefix(&cwd)
        .expect("scratch dirs live under the `.` preopen so `open-at` can reach them");

    let preopens = invoke("wasi:filesystem/preopens", "get-directories", vec![]);
    let root = first_preopen_descriptor(&preopens);

    // open-at(parent, path-flags=0, path, open-flags=OPEN_DIRECTORY, %flags=0)
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

    let result = types("[method]descriptor.read-directory", vec![root]);
    assert_read_directory_tuple(&result);

    assert_eq!(
        crate::stream_drain::read_directory(&dir),
        vec![
            ("regular-file".to_string(), "a.txt".to_string()),
            ("regular-file".to_string(), "b.txt".to_string()),
            ("directory".to_string(), "sub".to_string()),
        ]
    );
}

/// `read-directory: func() -> tuple<stream<directory-entry>,
///                                  future<result<_, error-code>>>`.
///
/// The SHAPE half of the assertion — that the answer is the declared tuple and
/// not the 0.2 `directory-entry-stream` resource. The ENTRIES are asserted
/// separately, through `stream_drain::read_directory`, because reading them
/// requires draining a `stream<directory-entry>` inside the same VM that opened
/// it and the one-host-call-per-`invoke` helper here cannot do that.
///
/// Both halves are needed and neither implies the other: this one catches a
/// host that goes back to answering a resource handle, and the drain catches a
/// host that answers the right shape with the wrong contents. For a while only
/// this half existed — a documented loss, now closed. The reason it could be
/// closed is `canon stream.read` learning element types; before that it copied
/// bytes and served `stream<u8>` and nothing else, so the honest position was
/// that the host answered the declared shape and nothing could read it — the
/// same blocker `tcp-socket.listen`'s `stream<tcp-socket>` sat behind.
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
fn read_directory_on_an_empty_directory_still_answers_the_tuple() {
    let dir = scratch_dir("read_directory_eof");
    let root = open_test_root(&dir);

    // An empty directory is not an error — it is an empty entry stream, and
    // the shape must be the same one a populated directory answers. 0.2 tested
    // this by reading `none` twice off the entry-stream resource; there is no
    // resource to read now.
    assert_read_directory_tuple(&types("[method]descriptor.read-directory", vec![root]));

    // And the drain terminates. This is the case that distinguishes a stream
    // the host CLOSED from one it merely has not written to: a closed empty
    // stream reads DROPPED and the loop ends, an unclosed one reads
    // COMPLETED-with-zero forever. Both are "no entries" to an assertion on
    // the result; only one of them returns.
    assert!(crate::stream_drain::read_directory(&dir).is_empty());
}

/// A fifo enumerates as `fifo`, and — more to the point — does not take the
/// rest of the directory down with it.
///
/// `variant descriptor-type` declares eight cases and the host answered
/// `"unknown"`, which is not one of them, for everything that was not a file,
/// directory or symlink. That is not a cosmetic mislabel: `canon stream.read`
/// lowers `%type` by looking the name up in the variant's case list and
/// writing the INDEX, so an undeclared name makes the copy fail and the
/// guest's whole drain error out. One fifo made every entry beside it
/// unreadable, and `/tmp` has fifos in it.
///
/// So this asserts the neighbour too. A host that reports the fifo as `other`
/// would still pass a test that only looked at the fifo.
#[test]
#[cfg(unix)]
fn read_directory_reports_a_fifo_without_failing_the_drain() {
    let dir = scratch_dir("read_directory_fifo");
    std::fs::write(dir.join("beside.txt"), "x").unwrap();
    let made = std::process::Command::new("mkfifo")
        .arg(dir.join("pipe"))
        .status()
        .expect("mkfifo should be runnable");
    assert!(made.success(), "mkfifo failed");

    assert_eq!(
        crate::stream_drain::read_directory(&dir),
        vec![
            ("regular-file".to_string(), "beside.txt".to_string()),
            ("fifo".to_string(), "pipe".to_string()),
        ]
    );
}

/// More entries than one `canon stream.read` asks for.
///
/// The guest lift requests `ENTRIES_PER_READ` (32) elements at a time and loops
/// until the copy reports anything but COMPLETED. Every other test in this file
/// fits in one read, so the loop-back — re-reading into the SAME buffer,
/// re-deriving `count`, re-testing the low nibble — never executes and the
/// whole second half of the drain is unproven.
///
/// 40 is deliberately just over the boundary. The failure it catches is silent
/// truncation at exactly 32, which no assertion on "did we get entries" would
/// notice and which any directory in real use hits before a user thinks to
/// check.
#[test]
fn read_directory_drains_more_entries_than_one_read_returns() {
    let dir = scratch_dir("read_directory_many");
    let mut expected: Vec<String> = Vec::new();
    for i in 0..40 {
        let name = format!("entry-{i:02}.txt");
        std::fs::write(dir.join(&name), "x").unwrap();
        expected.push(name);
    }
    expected.sort();

    assert_eq!(crate::stream_drain::read_directory_names(&dir), expected);
}

/// POSIX `readdir` yields `.` and `..`; `wasi:filesystem` does not.
///
/// Asserted rather than assumed because the difference is invisible until a
/// caller ported from `scandir` counts its entries, and because "the host
/// forgot to filter them" and "the host filtered them" produce the same
/// passing result on every other test in this file.
#[test]
fn read_directory_does_not_yield_dot_or_dotdot() {
    let dir = scratch_dir("read_directory_no_dots");
    std::fs::write(dir.join("only.txt"), "").unwrap();

    assert_eq!(
        crate::stream_drain::read_directory_names(&dir),
        vec!["only.txt".to_string()]
    );
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

// ── [method]descriptor.read-via-stream ────────────────────────────
//
// 0.3.1 answers `tuple<stream<u8>, future<result<_, error-code>>>` and the
// bytes come out of element 0 via `canon stream.read`. These used to drain
// through `wasi:io/streams.[method]input-stream.{read,blocking-read}` — a
// package 0.3.1 deleted — so they were asserting the 0.2 contract.

#[test]
fn read_via_stream_yields_the_file_bytes() {
    let dir = scratch_dir("read_stream");
    std::fs::write(dir.join("payload.bin"), b"hello, world").unwrap();
    let bytes = crate::stream_drain::read_via_stream(&dir, "payload.bin", 0.0);
    assert_eq!(bytes, b"hello, world");
}

#[test]
fn read_via_stream_returns_a_two_element_tuple() {
    let dir = scratch_dir("read_stream_tuple");
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
    let result = types(
        "[method]descriptor.read-via-stream",
        vec![descriptor, Value::F64(0.0)],
    );
    assert!(is_error(&result).is_none(), "read-via-stream should succeed");
    let Value::Object(object) = &result else {
        panic!("read-via-stream returns a tuple, got {result:?}");
    };
    let object = object.lock().unwrap();
    let ObjectKind::Array(parts) = &object.kind else {
        panic!("read-via-stream returns a tuple, got {:?}", object.kind);
    };
    assert_eq!(
        parts.len(),
        2,
        "tuple<stream<u8>, future<result<_, error-code>>> has exactly two elements"
    );
}

#[test]
fn read_via_stream_observes_the_offset() {
    let dir = scratch_dir("read_stream_offset");
    std::fs::write(dir.join("payload.bin"), b"abcdef").unwrap();
    // The offset positions the read like `pread`; the stream then carries
    // everything from there, so the suffix is what arrives — 0.2's separate
    // length argument has no counterpart in 0.3.1.
    let bytes = crate::stream_drain::read_via_stream(&dir, "payload.bin", 2.0);
    assert_eq!(bytes, b"cdef");
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

// `read_directory_rejects_bad_descriptor` below already covers what the old
// `directory_entry_stream_read_rejects_bad_descriptor` did: with the
// `directory-entry-stream` resource deleted, a bad handle can only be caught
// at `read-directory` itself, and that assertion already existed.

#[test]
fn read_directory_rejects_bad_descriptor() {
    let result = types("[method]descriptor.read-directory", vec![Value::Null]);
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

// Two `wasi:io/streams` `input-stream` tests stood here. The resource does not
// exist in 0.3.1 — a file is read through `read-via-stream`, which answers a
// `stream<u8>`. The bad-descriptor case they covered is still covered, by
// `read_via_stream_rejects_bad_descriptor` on the real verb.

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
        // `[method]directory-entry-stream.read-directory-entry` is NOT here:
        // 0.3.1 deleted the `directory-entry-stream` resource. `read-directory`
        // answers `tuple<stream<directory-entry>, future<...>>` directly.
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
