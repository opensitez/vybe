//! Go test helpers — parse, compile, and run programs through the Vybe VM.
//!
//! Batch macros: `go_run_cases!`, `go_compile_cases!`, `go_run_test!`, `go_compile_test!`,
//! plus aliases `compile_cases!` and `run_cases!` for feature-focused batches.
//! Compile-only stdlib batches (e.g. `test_net_http_compile.rs`, `test_os_exec_compile.rs`,
//! `test_log_flag_packages.rs`, `test_encoding_binary.rs`) use `go_compile_cases!` to assert the
//! frontend lowers imports,
//! types, and calls without requiring a live host implementation.
//! `test_go_statement_compile.rs` exercises distinct `go` statement forms (closures, methods,
//! parameterized literals, loop spawning) via `compile_cases!`.
//! `test_bytes_buffer_extended.rs` covers `bytes.Buffer` I/O semantics (Grow, ReadFrom, WriteTo,
//! UnreadByte/Rune, Next, Truncate, Reset, Equal) distinct from package-level `bytes.*` helpers
//! in `test_bytes_package.rs`.
//! Channel/select runtime batches (`test_channel_select_patterns_extra.rs`,
//! `test_select_patterns_advanced.rs`) use `go_run_cases!` for concrete stdout assertions and
//! `go_compile_cases!` for blocking or multi-case syntax the VM cannot execute synchronously.
//! `test_channel_buffered_patterns.rs` covers buffered `cap`/`len` runtime semantics, blocking
//! send/receive compile smoke, and fan-in merge patterns — distinct from
//! `test_channel_close_range.rs` (close, range, ok idioms).
//! `test_fmt_errors_print.rs` covers fmt Errorf message formatting, Sscanf/Fscanf parsing, and
//! Fprint/Fprintf/Fprintln writes into `bytes.Buffer` — distinct from `test_fmt_sprintf_verbs.rs`
//! (Sprintf verbs) and `test_errors_package.rs` (errors.Is / As / Unwrap).
//! `test_atomic_sync_extended.rs` covers `sync/atomic` Load, Store, Add, Swap, CompareAndSwap,
//! and Value — distinct from `test_sync_package.rs` (Mutex, RWMutex, WaitGroup, Once, Pool, Map).
//! `test_time_parse_format.rs` covers ParseDuration, Duration.String, custom Format/Parse layouts,
//! Add/AddDate/Sub, and Before/After/Equal — distinct from `test_time_package.rs` (RFC3339,
//! Unix epoch, Sleep/Tick/After compile smoke).
//! `test_regexp_advanced.rs` covers FindAllStringSubmatch, ReplaceAllString capture references,
//! LiteralPrefix, and NumSubexp runtime semantics — distinct from `test_regexp_package.rs`
//! (MatchString, FindStringSubmatch, basic ReplaceAllString, Split, compile-only NumSubexp).
//! `test_json_unmarshal_advanced.rs` covers advanced Unmarshal into struct/slice/map (embedded
//! fields, string tags, null pointer elements, unicode escapes), MarshalIndent layout/prefix,
//! and RawMessage envelope parsing — distinct from `test_json_marshal.rs` (basic Marshal/Unmarshal,
//! omitempty/tags, null, roundtrip, compile-only indent/RawMessage smoke).
//! `test_constants_iota_advanced.rs` covers advanced `iota` const blocks: flag OR/AND semantics,
//! offset/descending bit shifts, multi-blank skips, storage-unit ladders, and typed groups
//! (`int64`, `byte`, `uint32`, `rune`) — distinct from `test_iota_enumerations.rs` and
//! `test_constants.rs`.

#![allow(dead_code)]

use std::sync::{Arc, Mutex};
use vybe_runtime::{HostContext, VM, Value};

#[macro_export]
macro_rules! go_run_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            let out = $crate::helpers::run_prints($src);
            assert_eq!(out, $expected);
        }
    };
}

#[macro_export]
macro_rules! go_compile_test {
    ($name:ident, $src:expr) => {
        #[test]
        fn $name() {
            $crate::helpers::compile_ok($src);
        }
    };
}

#[macro_export]
macro_rules! go_run_cases {
    ($($name:ident => ($src:expr, $expected:expr),)+) => {
        $(go_run_test!($name, $src, $expected);)+
    };
}

#[macro_export]
macro_rules! go_compile_cases {
    ($($name:ident => $src:expr,)+) => {
        $(go_compile_test!($name, $src);)+
    };
}

/// Alias for `go_compile_cases!` — used by feature batches such as `test_go_statement_compile.rs`.
#[macro_export]
macro_rules! compile_cases {
    ($($name:ident => $src:expr,)+) => {
        $(go_compile_test!($name, $src);)+
    };
}

/// Alias for `go_run_cases!`.
#[macro_export]
macro_rules! run_cases {
    ($($name:ident => ($src:expr, $expected:expr),)+) => {
        $(go_run_test!($name, $src, $expected);)+
    };
}

fn compile_chunks(src: &str) -> Result<Vec<vybe_runtime::Chunk>, String> {
    {
        static R: std::sync::Once = std::sync::Once::new();
        R.call_once(vybe_language_go::register);
    }
    let module = vybe_language_go::parse(src)?;
    let profile = vybe_compiler::profile::parse_profile(vybe_language_go::profile_source())
        .map_err(|e| format!("profile parse failed: {}", e))?;
    vybe_compiler::primitives::Compiler::with_profile(profile).compile(&module)
}

pub fn compile_ok(src: &str) {
    match compile_chunks(src) {
        Ok(chunks) => {
            assert!(!chunks.is_empty(), "compile produced no chunks");
        }
        Err(e) => panic!("compile failed: {}", e),
    }
}

pub fn compile(src: &str) -> Vec<vybe_runtime::Chunk> {
    match compile_chunks(src) {
        Ok(c) => c,
        Err(e) => panic!("compile failed: {}", e),
    }
}

pub fn run(src: &str) -> Value {
    let chunks = compile(src);
    let mut vm = VM::new();
    vybe_compiler::primitives::platforms::init_platforms(&mut vm);
    vm.run(chunks).expect("run failed")
}

pub fn run_prints(src: &str) -> Vec<String> {
    let chunks = compile(src);
    let mut vm = VM::new();
    let output: Arc<Mutex<Vec<String>>> = Arc::new(Mutex::new(Vec::new()));
    let out = output.clone();
    vybe_compiler::primitives::platforms::init_platforms(&mut vm);
    vm.register_host_fn(
        "wasi:logging/logging",
        "log",
        Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
            let s: Vec<String> = args.iter().map(|a| format!("{}", a)).collect();
            out.lock().unwrap().push(s.join(" "));
            Value::Null
        }),
    );
    vybe_compiler::primitives::platforms::finalize_platforms(&mut vm);
    vm.run(chunks).expect("run failed");
    let result = output.lock().unwrap().clone();
    result
}

pub fn parse_ok(src: &str) -> bool {
    vybe_language_go::parse(src).is_ok()
}

pub fn compile_ok_check(src: &str) -> bool {
    compile_chunks(src).is_ok()
}

/// Run a `main` body that uses `fmt.Println`, wrapped in `package main`.
pub fn run_main_prints(body: &str) -> Vec<String> {
    run_prints(&format!(
        "package main; import \"fmt\"; func main() {{ {body} }}"
    ))
}
