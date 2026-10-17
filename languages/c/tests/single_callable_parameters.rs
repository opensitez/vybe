use std::path::{Path, PathBuf};
use std::process::Command;
use vybe_ast::{CallableInvocation, StmtKind};

const SOURCE: &str = include_str!("../../../tests/c/doom_startup/doom_intercept_single_callback.c");

fn fixture() -> PathBuf {
    Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../tests/c/doom_startup/doom_intercept_single_callback.c")
}

#[test]
fn callback_typedef_parameters_retain_single_callable_contract() {
    let module = vybe_language_c::parse(SOURCE).unwrap();
    assert_eq!(module.directives.callable_invocation, Some(CallableInvocation::Single));
    for name in ["P_TraverseIntercepts", "invoke_typedef", "invoke_explicit", "invoke_reassigned", "invoke_float"] {
        let params = module.body.iter().find_map(|statement| match &statement.kind {
            StmtKind::FunctionDecl { name: actual, params, .. } if actual == name => Some(params),
            _ => None,
        }).unwrap_or_else(|| panic!("missing {name}"));
        assert!(params[0].type_hint.as_deref().unwrap().starts_with("func("),
            "{name}: callback typedef was erased: {:?}", params[0].type_hint);
        if name == "invoke_float" {
            assert!(params[0].type_hint.as_deref().unwrap().ends_with("double"));
        }
    }
    let slots = vybe_language_c::parse(
        "typedef int (*callback_t)(int); int inspect_slots(callback_t *slot, callback_t table[2]) { return slot != 0 && table != 0; }",
    ).unwrap();
    let params = slots.body.iter().find_map(|statement| match &statement.kind {
        StmtKind::FunctionDecl { name, params, .. } if name == "inspect_slots" => Some(params),
        _ => None,
    }).unwrap();
    for param in params {
        assert!(!param.type_hint.as_deref().unwrap().starts_with("func("),
            "callback slot/array is not a scalar callable: {param:?}");
    }
}

struct NativeWorkspace(PathBuf);
impl Drop for NativeWorkspace {
    fn drop(&mut self) { let _ = std::fs::remove_dir_all(&self.0); }
}

#[test]
fn source_intercept_callback_has_static_comparisons_and_native_backing_parity() {
    let directory = NativeWorkspace(std::env::temp_dir()
        .join(format!("vybe-c-single-callback-{}", std::process::id())));
    std::fs::create_dir_all(&directory.0).unwrap();
    let executable = directory.0.join("native");
    let native = Command::new(std::env::var_os("CC").unwrap_or_else(|| "cc".into()))
        .args(["-std=c11", "-Wall", "-Wextra", "-Werror"])
        .arg(fixture()).arg("-o").arg(&executable).output().unwrap();
    assert!(native.status.success(), "{}", String::from_utf8_lossy(&native.stderr));
    let native = Command::new(&executable).output().unwrap();
    assert!(native.status.success(), "{}", String::from_utf8_lossy(&native.stderr));

    vybe_language_c::register();
    let chunks = vybe_compiler::projects::load(&fixture()).unwrap().compile().unwrap();
    for name in ["P_TraverseIntercepts", "invoke_typedef", "invoke_explicit", "invoke_reassigned", "invoke_float"] {
        let chunk = chunks.iter().find(|chunk| chunk.name == name).unwrap();
        let dump = vybe_runtime::debug::disassemble(chunk);
        assert!(!dump.contains("__vybe_dynlt"), "multicast handler loop in {name}: {dump}");
        assert!(dump.contains("call_ref"), "callback invocation missing in {name}: {dump}");
        if name == "P_TraverseIntercepts" {
            assert!(dump.contains("i32.lt_s"), "signed fixed_t field comparison: {dump}");
            assert!(dump.contains("i32.gt_s"), "signed fixed_t cutoff comparison: {dump}");
        }
    }
    let wasm = vybe_platform_wasm::write_wasm_with_options(&chunks,
        vybe_platform_wasm::WasmWriteOptions {
            include_vybe_metadata: false, include_name_section: true,
        });
    wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
        .validate_all(&wasm).unwrap();
    let mut vm = vybe_runtime::VM::new();
    vybe_runtime::init_all_registered(&mut vm, &vybe_runtime::capabilities::Capabilities::all());
    vm.run(chunks).unwrap();
    assert_eq!(vm.pending_exit_code, 0);
}
