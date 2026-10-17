//! Direct assertions separate memory behavior from snprintf output failures.
//! Source forms cover local arrays, record fields, heap bytes/words and literal
//! sources; backing selection is the frontend's responsibility, not inferred here.
use std::path::Path;

fn check(case: u8) {
    let path = Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../tests/c/doom_startup/memory_backing_matrix.c");
    let source = format!(
        "#define MEMORY_BACKING_CASE {case}\n{}",
        std::fs::read_to_string(&path).unwrap()
    );
    vybe_language_c::register();
    let mut vm = vybe_runtime::VM::new();
    vybe_runtime::init_all_registered(&mut vm, &vybe_runtime::capabilities::Capabilities::all());
    let compiled = vybe_compiler::dynamic::RuntimeCompilerService::new(&mut vm)
        .compile_source_by_name(source, "c", path)
        .unwrap();
    let source_functions = [
        "main",
        "copy_content",
        "ordered_overlap",
        "zero_and_fill",
        "byte_search",
        "realloc_preservation",
    ];
    let mut equality_failures = Vec::new();
    for chunk in &compiled.chunks {
        if source_functions.contains(&chunk.name.as_str()) {
            let dump = vybe_runtime::debug::disassemble(chunk);
            for forbidden in ["dyneq", "abstractEqNe"] {
                if dump.contains(forbidden) {
                    equality_failures.push(format!("{}: {forbidden}", chunk.name));
                }
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
    let validation = wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
        .validate_all(&wasm)
        .map(|_| ())
        .map_err(|error| error.to_string());
    let execution = vybe_compiler::dynamic::RuntimeCompilerService::new(&mut vm)
        .run_compiled(compiled)
        .map(|_| vm.pending_exit_code)
        .map_err(|error| error.to_string());
    eprintln!(
        "memory case {case}: validation={validation:?}; execution={execution:?}; \
        static-equality audit={equality_failures:?}"
    );
    assert!(validation.is_ok(), "case {case}: WASM: {validation:?}");
    assert_eq!(execution, Ok(0), "case {case}: source");
    assert!(
        equality_failures.is_empty(),
        "case {case}: {equality_failures:?}"
    );
}

#[test]
fn memchr_preserves_interior_views_and_unsigned_bytes() {
    let path = Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../tests/c/doom_startup/memory_memchr_views.c");
    vybe_language_c::register();
    let mut vm = vybe_runtime::VM::new();
    vybe_runtime::init_all_registered(&mut vm, &vybe_runtime::capabilities::Capabilities::all());
    let compiled = vybe_compiler::dynamic::RuntimeCompilerService::new(&mut vm)
        .compile_path(&path)
        .unwrap();
    let search = compiled
        .chunks
        .iter()
        .find(|chunk| chunk.name == "__c_mem_index_of_h")
        .expect("byte search helper emitted");
    let mut offset = 0;
    let mut instructions = 0;
    while offset < search.code.len() {
        let (_, next) = vybe_runtime::debug::disassemble_instruction(search, offset);
        assert!(next > offset, "instruction decoder must advance");
        offset = next;
        instructions += 1;
    }
    eprintln!(
        "memchr helper: {instructions} emitted instructions, {} encoded bytes",
        search.code.len()
    );
    let equality_failures: Vec<_> = compiled
        .chunks
        .iter()
        .filter(|chunk| matches!(chunk.name.as_str(), "main" | "check_bytes"))
        .filter_map(|chunk| {
            let dump = vybe_runtime::debug::disassemble(chunk);
            let lines: Vec<_> = dump.lines().collect();
            for (index, line) in lines.iter().enumerate() {
                if line.contains("dyneq") || line.contains("abstractEqNe") {
                    eprintln!(
                        "{} equality site:\n{}",
                        chunk.name,
                        lines[index.saturating_sub(8)..(index + 4).min(lines.len())].join("\n")
                    );
                }
            }
            (dump.contains("dyneq") || dump.contains("abstractEqNe")).then_some(chunk.name.clone())
        })
        .collect();
    let wasm = vybe_platform_wasm::write_wasm_with_options(
        &compiled.chunks,
        vybe_platform_wasm::WasmWriteOptions {
            include_vybe_metadata: false,
            include_name_section: true,
        },
    );
    let validation = wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
        .validate_all(&wasm)
        .map(|_| ())
        .map_err(|error| error.to_string());
    vm.record_runtime_perf(true);
    let started = std::time::Instant::now();
    let execution = vybe_compiler::dynamic::RuntimeCompilerService::new(&mut vm)
        .run_compiled(compiled)
        .map(|_| vm.pending_exit_code)
        .map_err(|error| error.to_string());
    let elapsed = started.elapsed();
    let counters = vm.take_runtime_perf().expect("requested debug counters");
    eprintln!(
        "memchr views: validation={validation:?}; execution={execution:?}; noEq={equality_failures:?}; \
         debug dispatches={}; load8_u={}; elapsed={elapsed:?}",
        counters.dispatches, counters.i32_load8_u
    );
    assert!(validation.is_ok(), "WASM: {validation:?}");
    assert_eq!(execution, Ok(0));
    assert!(equality_failures.is_empty(), "{equality_failures:?}");
}

#[test]
fn memchr_linear_zero_count_and_access_trap() {
    // Guest/WASM boundary checks, deliberately not a native-C conformance test:
    // a nonzero search at this invalid linear address must retain its load trap.
    for count in [0, 1] {
        vybe_language_c::register();
        let mut vm = vybe_runtime::VM::new();
        vybe_runtime::init_all_registered(
            &mut vm,
            &vybe_runtime::capabilities::Capabilities::all(),
        );
        let source = format!(
            "#include <string.h>\nint main(void) {{ memchr((void*)0xffffffff, 17, {count}); return 0; }}"
        );
        let compiled = vybe_compiler::dynamic::RuntimeCompilerService::new(&mut vm)
            .compile_source_by_name(
                source,
                "c",
                Path::new(env!("CARGO_MANIFEST_DIR")).join("memchr_linear_boundary.c"),
            )
            .unwrap();
        assert!(
            compiled
                .chunks
                .iter()
                .any(|chunk| vybe_runtime::debug::disassemble(chunk).contains("i32.load8_u")),
            "search must use the existing standard linear byte load"
        );
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
        let execution =
            vybe_compiler::dynamic::RuntimeCompilerService::new(&mut vm).run_compiled(compiled);
        if count == 0 {
            execution.expect("zero count must not issue a memory access");
            assert_eq!(vm.pending_exit_code, 0);
        } else {
            let error = execution
                .expect_err("invalid nonzero search must trap")
                .to_lowercase();
            assert!(error.contains("out of bounds"), "unexpected trap: {error}");
        }
    }
}

#[test]
fn copy_content_and_literal_bridge() {
    check(1);
}
#[test]
fn bcopy_and_memmove_overlap() {
    check(2);
}
#[test]
fn bzero_and_fill_record_fields() {
    check(3);
}
#[test]
fn memchr_byte_identity() {
    check(4);
}
#[test]
fn realloc_preserves_bytes_and_words() {
    check(5);
}
