use std::path::Path;
use vybe_runtime::VM;

fn matches_native(fixture: &str, no_dynamic_equality: bool) {
    vybe_language_c::register();
    let path = Path::new(env!("CARGO_MANIFEST_DIR")).join(format!("../../tests/c/{fixture}.c"));
    let native = Path::new(env!("CARGO_MANIFEST_DIR")).join(format!(
        "../../target/debug/{}-{}-native",
        path.file_stem().unwrap().to_str().unwrap(),
        std::process::id(),
    ));
    assert!(
        std::process::Command::new("cc")
            .arg("-O0")
            .arg(&path)
            .arg("-o")
            .arg(&native)
            .status()
            .unwrap()
            .success(),
        "{fixture}: native compilation"
    );
    assert!(
        std::process::Command::new(&native)
            .status()
            .unwrap()
            .success(),
        "{fixture}: native execution"
    );
    let mut vm = VM::new();
    vybe_runtime::init_all_registered(&mut vm, &vybe_runtime::capabilities::Capabilities::all());
    let compiled = vybe_compiler::dynamic::RuntimeCompilerService::new(&mut vm)
        .compile_path(&path)
        .unwrap();
    if let Some(dir) = std::env::var_os("C_POINTER_WASM_DIR") {
        for chunk in &compiled.chunks {
            std::fs::write(
                Path::new(&dir).join(format!(
                    "{}-{}.dump",
                    path.file_stem().unwrap().to_str().unwrap(),
                    chunk.name
                )),
                vybe_runtime::debug::disassemble(chunk),
            )
            .unwrap();
        }
    }
    if no_dynamic_equality {
        for chunk in &compiled.chunks {
            let dump = vybe_runtime::debug::disassemble(chunk);
            assert!(
                !dump.contains("dyneq"),
                "{fixture}/{}: dynamic equality",
                chunk.name
            );
            let imports = if chunk.imports.is_empty() {
                &compiled.chunks[0].imports
            } else {
                &chunk.imports
            };
            let mut offset = 0;
            while offset + 4 <= chunk.code.len() {
                use vybe_runtime::opcode::Op;
                let op = Op::decode(chunk.read_u16(offset), chunk.read_u16(offset + 2)).unwrap();
                if op == Op::CALL {
                    let index = chunk.read_u16(offset + 4) as usize;
                    if let Some(import) = imports.get(index) {
                        assert!(
                            import.module != "ecma:value"
                                || !matches!(import.name.as_str(), "abstractEq" | "abstractNe"),
                            "{fixture}/{}: dynamic equality call at {offset}",
                            chunk.name
                        );
                    }
                }
                let (_, next) = vybe_runtime::debug::disassemble_instruction(chunk, offset);
                assert!(next > offset);
                offset = next;
            }
        }
    }
    let wasm = vybe_platform_wasm::write_wasm_with_options(
        &compiled.chunks,
        vybe_platform_wasm::WasmWriteOptions {
            include_vybe_metadata: false,
            include_name_section: true,
        },
    );
    wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
        .validate_all(&wasm)
        .unwrap();
    if let Some(dir) = std::env::var_os("C_POINTER_WASM_DIR") {
        std::fs::write(
            Path::new(&dir).join(format!(
                "{}.wasm",
                path.file_stem().unwrap().to_str().unwrap()
            )),
            wasm,
        )
        .unwrap();
    }
    let result =
        vybe_compiler::dynamic::RuntimeCompilerService::new(&mut vm).run_compiled(compiled);
    if fixture == "doom_startup/pointer_byte_distance_contracts" {
        for name in ["managed_byte_distance_result", "heap_byte_distance_result"] {
            let value = vm.global(name).unwrap();
            assert_eq!(value.as_i32(), 24, "{name}: {value:?}");
        }
    }
    result.unwrap_or_else(|error| panic!("{fixture}: {error}"));
    assert_eq!(vm.pending_exit_code, 0, "{fixture}");
    eprintln!("{fixture}: PASS");
}

#[test]
fn pointer_to_pointer_basic_matches_native() {
    matches_native("pointer_advanced/pointer_to_pointer_basic", false);
}

#[test]
fn pointer_to_pointer_write_through_matches_native() {
    matches_native("pointer_advanced/pointer_to_pointer_write_through", false);
}

#[test]
fn pointer_to_pointer_parameters_and_reassignment_match_native() {
    let mut failures = Vec::new();
    for fixture in [
        "parameter_passing/pointer_to_pointer_parameter_can_follow_two_levels",
        "functions_advanced/function_can_consume_pointer_to_pointer",
        "pointers_basic/pointer_to_pointer_can_follow_reassigned_pointer",
        "pointers_basic/pointer_to_pointer_can_update_original_value",
    ] {
        if std::panic::catch_unwind(std::panic::AssertUnwindSafe(|| {
            matches_native(fixture, false);
        }))
        .is_err()
        {
            failures.push(fixture);
        }
    }
    assert!(
        failures.is_empty(),
        "pointer-to-pointer fixtures failed: {failures:?}"
    );
}

#[test]
fn pointer_stride_matrix_matches_native_without_dynamic_equality() {
    matches_native("doom_startup/pointer_stride_matrix", true);
}

#[test]
fn pointer_stride_matrix_execution_matches_native() {
    matches_native("doom_startup/pointer_stride_matrix", false);
}

#[test]
fn pointer_array_cursor_contracts_match_native() {
    matches_native("doom_startup/pointer_array_cursor_contracts", false);
}

#[test]
fn array_of_pointers_increment_matches_native() {
    matches_native(
        "lang_pointer_arithmetic_steps/array_of_pointers_increment_selects_next",
        false,
    );
}

#[test]
fn pointer_difference_formatting_matches_native() {
    for fixture in [
        "int_pointer_subtract_yields_element_count",
        "char_ptrdiff_counts_bytes",
        "ptrdiff_same_element_zero",
        "ptrdiff_both_one_past_end_is_zero",
        "ptrdiff_one_past_end_minus_start",
        "ptrdiff_one_past_end_minus_one_before",
        "ptrdiff_signed_int_array_span",
        "ptrdiff_after_increment_chain",
        "ptrdiff_after_decrement_chain",
    ] {
        matches_native(&format!("lang_pointer_arithmetic_steps/{fixture}"), false);
    }
}

#[test]
fn record_pointer_byte_difference_matches_native() {
    matches_native("doom_startup/pointer_byte_distance_contracts", false);
}
