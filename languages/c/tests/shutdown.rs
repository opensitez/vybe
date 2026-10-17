use std::sync::{Arc, Mutex};

fn run_fixture(name: &str) -> (i32, String) {
    let (status, output, _) = run_fixture_with_write_count(name);
    (status, output)
}

fn run_fixture_with_write_count(name: &str) -> (i32, String, usize) {
    vybe_language_c::register();
    let path = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("tests/fixtures")
        .join(name);
    let chunks = vybe_compiler::projects::load(&path)
        .unwrap()
        .compile()
        .unwrap();
    let bytes = vybe_platform_wasm::write_wasm_with_options(
        &chunks,
        vybe_platform_wasm::WasmWriteOptions {
            include_vybe_metadata: false,
            include_name_section: true,
        },
    );
    wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
        .validate_all(&bytes)
        .unwrap();
    let output = Arc::new(Mutex::new(Vec::new()));
    let captured = output.clone();
    let writes = Arc::new(std::sync::atomic::AtomicUsize::new(0));
    let captured_writes = writes.clone();
    let mut vm = vybe_runtime::VM::new();
    vybe_runtime::init_all_registered(&mut vm, &vybe_runtime::capabilities::Capabilities::all());
    vm.register_host_fn(
        "wasi:cli/stdout",
        "write-via-stream",
        Box::new(move |ctx, args| {
            captured_writes.fetch_add(1, std::sync::atomic::Ordering::Relaxed);
            captured.lock().unwrap().extend(ctx.stream_drain(&args[0]));
            let (future, id) = ctx.create_future();
            ctx.resolve_future(id, vybe_runtime::Value::Null);
            future
        }),
    );
    vm.run(chunks).unwrap();
    let text = String::from_utf8(output.lock().unwrap().clone()).unwrap();
    (
        vm.pending_exit_code,
        text,
        writes.load(std::sync::atomic::Ordering::Relaxed),
    )
}

#[test]
fn silent_completion_does_not_open_an_output_stream() {
    assert_eq!(
        run_fixture_with_write_count("shutdown_empty.c"),
        (0, String::new(), 0),
    );
}

#[test]
fn normal_completion_captures_locals_and_drains_live_lifo_registry() {
    assert_eq!(run_fixture("shutdown_normal.c"), (7, "PBL7:42AAA".into()));
}

#[test]
fn nested_exit_runs_cleanup_and_flushes_without_newline() {
    assert_eq!(run_fixture("shutdown_nested.c"), (7, "PH".into()));
}

#[test]
fn reentrant_exit_does_not_replay_callbacks_or_resume_outer_cleanup() {
    assert_eq!(run_fixture("shutdown_reentrant.c"), (9, "PH".into()));
}

#[test]
fn quick_exit_uses_only_quick_registry_including_new_registrations() {
    assert_eq!(run_fixture("shutdown_quick.c"), (3, "Q\nL\n".into()));
}

#[test]
fn immediate_exit_bypasses_callbacks_and_partial_output() {
    assert_eq!(run_fixture("shutdown_immediate.c"), (4, "".into()));
}

#[test]
fn expression_exit_is_nonreturning_and_runs_cleanup() {
    assert_eq!(run_fixture("shutdown_expression.c"), (6, "PH".into()));
}

#[test]
fn main_fallthrough_runs_cleanup_and_flushes_partial_output() {
    assert_eq!(run_fixture("shutdown_fallthrough.c"), (0, "PH".into()));
}

#[test]
fn abort_raises_sigabrt_then_bypasses_cleanup_if_handler_returns() {
    assert_eq!(run_fixture("shutdown_abort.c"), (134, "signal\n".into()));
}

#[test]
fn abort_handler_can_exit_normally_with_cleanup_and_its_own_status() {
    assert_eq!(
        run_fixture("shutdown_abort_handler_exit.c"),
        (8, "PH".into())
    );
}

#[test]
fn expression_abort_is_nonreturning_even_when_sigabrt_is_ignored() {
    assert_eq!(run_fixture("shutdown_abort_expression.c"), (134, "".into()));
}

#[test]
fn default_abort_bypasses_callbacks_and_partial_output() {
    assert_eq!(run_fixture("shutdown_abort_default.c"), (134, "".into()));
}

#[test]
fn direct_void_pointer_call_and_relay_preserve_managed_and_linear_addresses() {
    assert_eq!(
        run_fixture("shutdown_void_pointer_control.c"),
        (0, "7:42,7:55,7:66,7:77".into())
    );
}

#[test]
fn on_exit_preserves_live_scalar_record_and_linear_pointer_backings() {
    assert_eq!(
        run_fixture("shutdown_pointer_backings.c"),
        (0, "66,55,88,".into())
    );
}
