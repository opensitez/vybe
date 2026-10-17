use std::path::{Path, PathBuf};
use std::process::Command;
use std::time::Instant;

use vybe_runtime::{Chunk, opcode::Op};

struct NativeWorkspace(PathBuf);

impl Drop for NativeWorkspace {
    fn drop(&mut self) {
        let _ = std::fs::remove_dir_all(&self.0);
    }
}

fn check_native(source: &Path, case: &str) {
    let directory = NativeWorkspace(
        std::env::temp_dir().join(format!("vybe-doom-pattern-{}-{case}", std::process::id())),
    );
    std::fs::create_dir_all(&directory.0).unwrap();
    let executable = directory.0.join("reference");
    let compiler = std::env::var_os("CC").unwrap_or_else(|| "cc".into());
    let compilation = Command::new(compiler)
        .args(["-std=gnu11", "-O0"])
        .arg(source)
        .arg("-o")
        .arg(&executable)
        .output()
        .expect("native C reference compiler (CC or cc) is required");
    assert!(
        compilation.status.success(),
        "{case}: native compilation failed:\n{}",
        String::from_utf8_lossy(&compilation.stderr)
    );
    let execution = Command::new(executable).output().unwrap();
    assert!(
        execution.status.success(),
        "{case}: native assertions failed ({:?}):\n{}\n{}",
        execution.status.code(),
        String::from_utf8_lossy(&execution.stdout),
        String::from_utf8_lossy(&execution.stderr)
    );
}

fn static_equality_errors(chunks: &[Chunk]) -> Vec<String> {
    let mut errors = Vec::new();
    for chunk in chunks {
        if vybe_runtime::debug::disassemble(chunk).contains("dyneq") {
            errors.push(format!("{}: dynamic equality", chunk.name));
        }
        let imports = if chunk.imports.is_empty() {
            &chunks[0].imports
        } else {
            &chunk.imports
        };
        let mut offset = 0;
        while offset + 4 <= chunk.code.len() {
            let op = Op::decode(chunk.read_u16(offset), chunk.read_u16(offset + 2)).unwrap();
            if op == Op::CALL {
                if let Some(import) = imports.get(chunk.read_u16(offset + 4) as usize) {
                    if import.module == "ecma:value"
                        && matches!(import.name.as_str(), "abstractEq" | "abstractNe")
                    {
                        errors.push(format!(
                            "{}@{offset}: dynamic equality host call",
                            chunk.name
                        ));
                    }
                }
            }
            let (_, next) = vybe_runtime::debug::disassemble_instruction(chunk, offset);
            assert!(next > offset, "{}: decoder did not advance", chunk.name);
            offset = next;
        }
    }
    errors
}

fn check_pattern(case: &str, fixture: &str) {
    let source = Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../tests/c/doom_startup")
        .join(fixture);
    check_native(&source, case);
    vybe_language_c::register();
    let started = Instant::now();
    let chunks = vybe_compiler::projects::load(&source)
        .unwrap()
        .compile()
        .unwrap();
    let compile = started.elapsed();
    let mut errors = static_equality_errors(&chunks);
    let wasm = vybe_platform_wasm::write_wasm_with_options(
        &chunks,
        vybe_platform_wasm::WasmWriteOptions {
            include_vybe_metadata: false,
            include_name_section: true,
        },
    );
    wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
        .validate_all(&wasm)
        .unwrap_or_else(|error| panic!("{case}: invalid WASM: {error}"));
    let mut vm = vybe_runtime::VM::new();
    vybe_runtime::init_all_registered(&mut vm, &vybe_runtime::capabilities::Capabilities::all());
    let started = Instant::now();
    if let Err(error) = vm.run(chunks) {
        errors.push(format!("VM: {error}"));
    }
    if vm.pending_exit_code != 0 {
        errors.push(format!("C assertion exit status: {}", vm.pending_exit_code));
    }
    assert!(errors.is_empty(), "{case}:\n{}", errors.join("\n"));
    eprintln!(
        "{case}: native/WASM/static-equality/VM PASS; compile={compile:?}, execution={:?}",
        started.elapsed()
    );
}

macro_rules! patterns {
    ($($case:ident => $fixture:literal),+ $(,)?) => {
        $(#[test] fn $case() { check_pattern(stringify!($case), $fixture); })+
    };
}

patterns! {
    pointer_short_stride => "linear_short_pointer_offsets.c",
    pointer_array_and_record_stride => "pointer_stride_matrix.c",
    pointer_nested_list_backlinks => "nested_pointer_list_insert.c",
    pointer_table_single_evaluation => "pointer_table_single_evaluation.c",
    pointer_record_slot_receiver => "record_pointer_slot_receiver.c",
    pointer_record_null_aliases => "pointer_record_null_aliases.c",
    pointer_branch_join_backings => "pointer_branch_join_backings.c",
    pointer_erased_first_member_container => "void_first_member_container.c",
    pointer_first_member_location_identity => "first_member_pointer_identity.c",
    pointer_first_member_union_write_through => "first_member_union_update.c",
    pointer_embedded_first_member_roundtrip => "embedded_first_member_address.c",
    pointer_sizeof_indexed_record_allocation => "sizeof_indexed_record_allocation.c",
    memory_static_record_size_reset => "static_record_size_reset.c",
    memory_qualified_aggregate_layout => "qualified_aggregate_layout.c",
    memory_short_integer_spellings => "short_integer_spellings.c",
    pointer_word_byte_alias => "byte_view_of_word_array.c",
    memory_copy_fill_overlap_search_reallocate => "memory_backing_matrix.c",
    memory_forwarded_conditional_write => "conditional_buffer_write_through_forwarded_pointer.c",
    memory_pointer_results_and_single_evaluation => "memory_return_pointer_types.c",
    memory_unsigned_byte_copy_result_identity => "unsigned_byte_memcpy_return_identity.c",
    memory_record_zero_fill_preserves_aliases => "managed_record_memset_reset.c",
    memory_narrow_string_bytes_and_utf8 => "narrow_string_escape_bytes.c",
    update_scalar_widths_and_single_evaluation => "scalar_pointer_integer_updates.c",
    update_hud_alias => "hud_ammo_pointer_alias.c",
    update_pistol_indexed_compound => "doom_pistol_ammo_update.c",
    arithmetic_unsigned_wrap_and_promotions => "unsigned_word_arithmetic.c",
    arithmetic_signed_byte_postincrement => "doom_demo_signed_postincrement.c",
    condition_negation_and_short_circuit => "negated_static_conditions.c",
    arithmetic_fixed_point_wide_intermediate => "fixed_point_wide_intermediate.c",
    arithmetic_unsigned_texture_coordinate_wrap => "unsigned_span_position_wrap.c",
    input_keyboard_command_bits => "keyboard_command_input.c",
    input_nested_tic_command_copy => "tic_command_transfer.c",
    pickup_radius_height_health_and_record_backings => "doom_pickup_contact.c",
    renderer_wall_plane_spans => "doom_wall_plane_spans.c",
}
