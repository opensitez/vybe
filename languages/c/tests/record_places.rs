use std::path::Path;

fn assert_no_dynamic_equality(chunks: &[vybe_runtime::Chunk]) {
    use vybe_runtime::opcode::Op;
    for chunk in chunks {
        let dump = vybe_runtime::debug::disassemble(chunk);
        if dump.contains("dyneq") {
            let lines: Vec<_> = dump.lines().collect();
            let at = lines.iter().position(|line| line.contains("dyneq")).unwrap();
            panic!("dynamic equality helper in {}:\n{}", chunk.name,
                lines[at.saturating_sub(8)..(at + 8).min(lines.len())].join("\n"));
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
                let index = chunk.read_u16(offset + 4) as usize;
                if let Some(import) = imports.get(index) {
                    assert!(import.module != "ecma:value"
                        || !matches!(import.name.as_str(), "abstractEq" | "abstractNe"),
                        "dynamic equality host call in {} at {offset}: {}:{}",
                        chunk.name, import.module, import.name);
                }
            }
            let (_, next) = vybe_runtime::debug::disassemble_instruction(chunk, offset);
            assert!(next > offset, "decoder did not advance in {}", chunk.name);
            offset = next;
        }
    }
}

#[test]
fn glow_nested_record_fields_match_native_and_emit_valid_static_wasm() {
    assert_native_static_record_fixture("doom_glow_kernel.c", Some("T_Glow"));
}

#[test]
fn nested_record_captures_preserve_owner_across_recursive_rhs() {
    assert_native_static_record_fixture("nested_record_capture_reentrant.c", None);
}

#[test]
fn pointer_field_captures_preserve_target_across_recursive_rhs() {
    assert_native_static_record_fixture("reentrant_pointer_field_assignment.c", None);
}

#[test]
fn sky_planes_draw_expected_pixels_through_real_column_callback() {
    assert_native_static_record_fixture("doom_sky_plane_pixels.c", None);
}

#[test]
fn real_wall_clipping_matches_per_column_visibility() {
    assert_native_static_record_fixture("doom_wall_clip_visibility.c", None);
}

#[test]
fn indexed_sector_pointers_transfer_through_real_doom_record_layouts() {
    assert_native_static_record_fixture("indexed_sector_pointer_transfer.c", None);
}

#[test]
fn real_map_loaders_preserve_sector_links_in_zone_allocated_records() {
    assert_native_static_record_fixture("doom_map_sector_links.c", None);
}

#[test]
fn real_map_loaders_preserve_sector_links_on_guest_worker() {
    assert_native_static_record_fixture_with_flags(
        "doom_map_sector_links_threaded.c", None, &["-pthread"],
    );
}

#[test]
fn real_wad_cache_hits_and_mapped_lumps_preserve_tags_and_aliases() {
    assert_native_static_record_fixture("doom_wad_cache_paths.c", Some("W_CacheLumpNum"));
}

#[test]
fn record_reads_preserve_direct_view_and_linear_aliases_after_mutation() {
    assert_native_static_record_fixture("record_read_backing_once.c", None);
}

#[test]
fn embedded_record_zero_index_preserves_singleton_alias() {
    assert_native_static_record_fixture("embedded_record_zero_index.c", None);
}

#[test]
fn nested_pointer_conditions_preserve_c_truth_and_short_circuiting() {
    assert_native_static_record_fixture("nested_pointer_boolean_conditions.c", None);
}

#[test]
fn unsigned_record_switch_selectors_preserve_all_sixty_four_bits() {
    assert_native_static_record_fixture("unsigned_record_switch_selector.c", None);
}

#[test]
fn initialized_header_pointers_preserve_backing_and_single_evaluation() {
    assert_native_static_record_fixture("initialized_header_pointer_backing.c", None);
}

#[test]
fn signed_field_word_extensions_preserve_values_and_single_evaluation() {
    assert_native_static_record_fixture("signed_field_word_extension.c", None);
}

#[test]
fn scalar_array_memcmp_preserves_byte_layout_and_argument_effects() {
    assert_native_static_record_fixture("memcmp_scalar_array_layout.c", None);
}

#[test]
fn comparison_views_preserve_offsets_nuls_mutations_and_argument_effects() {
    assert_native_static_record_fixture("string_compare_views.c", None);
}

#[test]
fn real_zone_cache_purge_release_and_level_restart_clear_owners() {
    assert_native_static_record_fixture("doom_zone_cache_pressure.c", None);
}

#[test]
fn real_melt_columns_match_pixels_and_cursor_state_each_frame() {
    let path = Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../tests/c/doom_startup/doom_melt_pixels.c");
    vybe_language_c::register();
    let module = vybe_compiler::projects::load(&path).unwrap().prepared_module().unwrap();
    let body = module.body.iter().find_map(|statement| match &statement.kind {
        vybe_ast::StmtKind::FunctionDecl { name, body, .. } if name == "wipe_doMelt" => Some(body),
        _ => None,
    }).unwrap();
    let mut transfers = 0;
    for statement in body {
        statement.walk_exprs(&mut |value| {
            if let vybe_ast::ExprKind::Call { callee, .. } = &value.kind {
                transfers += usize::from(matches!(&callee.kind,
                    vybe_ast::ExprKind::Ident(name) if name == "__c_ptr_copy_column16"));
            }
        });
    }
    assert_eq!(transfers, 2, "both real column loops must retain guarded word transfers");
    assert_native_static_record_fixture("doom_melt_pixels.c", Some("wipe_doMelt"));
}

fn assert_native_static_record_fixture(filename: &str, report_body: Option<&str>) {
    assert_native_static_record_fixture_with_flags(filename, report_body, &[]);
}

#[test]
fn directory_stream_read_and_drop_keep_declared_entry_type() {
    assert_native_static_record_fixture("directory_typed_stream_drop.c", None);
}

#[test]
fn directory_close_comparisons_are_static_and_evaluate_arguments_once() {
    assert_native_static_record_fixture("directory_close_static_comparison.c", None);
}

fn assert_native_static_record_fixture_with_flags(
    filename: &str, report_body: Option<&str>, native_flags: &[&str],
) {
    let source = Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../tests/c/doom_startup").join(filename);
    assert_native_map_fixture_with_flags(&source, native_flags);
    vybe_language_c::register();
    let mut vm = vybe_runtime::VM::new();
    vybe_runtime::init_all_registered(&mut vm, &vybe_runtime::capabilities::Capabilities::all());
    let compiled = vybe_compiler::dynamic::RuntimeCompilerService::new(&mut vm)
        .compile_path(&source).unwrap();
    assert_no_dynamic_equality(&compiled.chunks);
    if let Some(name) = report_body {
        let chunk = compiled.chunks.iter().find(|chunk| chunk.name == name).unwrap();
        eprintln!("{name} bytecode baseline: {} bytes", chunk.code.len());
    }
    let bytes = vybe_platform_wasm::write_wasm_with_options(&compiled.chunks,
        vybe_platform_wasm::WasmWriteOptions {
            include_vybe_metadata: false,
            include_name_section: true,
        });
    validate_wasm_with_context(&bytes, &compiled.chunks);
    vybe_compiler::dynamic::RuntimeCompilerService::new(&mut vm)
        .run_compiled(compiled).unwrap();
    assert_eq!(vm.pending_exit_code, 0);
}

#[test]
fn endoom_links_the_real_sdl_textscreen_functions() {
    let path = Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../tests/c/doom_startup/endoom_textscreen_link.c");
    vybe_language_c::register();
    let mut vm = vybe_runtime::VM::new();
    vybe_runtime::init_all_registered(&mut vm, &vybe_runtime::capabilities::Capabilities::all());
    let compiled = vybe_compiler::dynamic::RuntimeCompilerService::new(&mut vm)
        .compile_path(&path).unwrap();
    for name in ["TXT_Init", "TXT_GetScreenData", "TXT_UpdateScreen", "TXT_GetChar",
        "TXT_Sleep", "TXT_Shutdown", "I_Endoom"] {
        assert!(compiled.chunks.iter().any(|chunk| chunk.name == name),
            "missing textscreen implementation: {name}");
    }
    let bytes = vybe_platform_wasm::write_wasm_with_options(&compiled.chunks,
        vybe_platform_wasm::WasmWriteOptions {
            include_vybe_metadata: false,
            include_name_section: true,
        });
    validate_wasm_with_context(&bytes, &compiled.chunks);
    vybe_compiler::dynamic::RuntimeCompilerService::new(&mut vm)
        .run_compiled(compiled).unwrap();
    let callable = vm.global("check_textscreen_link").expect("textscreen link check").clone();
    assert_eq!(vm.invoke(&callable, &[]).unwrap().as_i32(), 73);
}

#[test]
fn doom_collision_sides_match_integer_cross_products() {
    let path = Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../tests/c/doom_startup/doom_collision_sides.c");
    vybe_language_c::register();
    let mut vm = vybe_runtime::VM::new();
    vybe_runtime::init_all_registered(&mut vm, &vybe_runtime::capabilities::Capabilities::all());
    let compiled = vybe_compiler::dynamic::RuntimeCompilerService::new(&mut vm)
        .compile_path(&path).unwrap();
    assert_no_dynamic_equality(&compiled.chunks);
    let bytes = vybe_platform_wasm::write_wasm_with_options(&compiled.chunks,
        vybe_platform_wasm::WasmWriteOptions {
            include_vybe_metadata: false,
            include_name_section: true,
        });
    validate_wasm_with_context(&bytes, &compiled.chunks);
    vybe_compiler::dynamic::RuntimeCompilerService::new(&mut vm)
        .run_compiled(compiled).unwrap();
    assert_eq!(vm.pending_exit_code, 0);
}

#[test]
fn sdl_tick_comparison_uses_declared_unsigned_word_type() {
    let path = Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../tests/c/doom_startup/sdl_tick_comparison.c");
    vybe_language_c::register();
    let mut vm = vybe_runtime::VM::new();
    vybe_runtime::init_all_registered(&mut vm, &vybe_runtime::capabilities::Capabilities::all());
    let compiled = vybe_compiler::dynamic::RuntimeCompilerService::new(&mut vm)
        .compile_path(&path).unwrap();
    assert_no_dynamic_equality(&compiled.chunks);
    let bytes = vybe_platform_wasm::write_wasm_with_options(&compiled.chunks,
        vybe_platform_wasm::WasmWriteOptions {
            include_vybe_metadata: false,
            include_name_section: true,
        });
    validate_wasm_with_context(&bytes, &compiled.chunks);
    vybe_compiler::dynamic::RuntimeCompilerService::new(&mut vm)
        .run_compiled(compiled).unwrap();
    assert_eq!(vm.pending_exit_code, 0);
}

#[test]
fn doom_visplanes_preserve_inline_array_fill_after_pointer_reassignment() {
    let path = Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../tests/c/doom_startup/doom_visplane_reuse.c");
    vybe_language_c::register();
    let mut vm = vybe_runtime::VM::new();
    vybe_runtime::init_all_registered(&mut vm, &vybe_runtime::capabilities::Capabilities::all());
    let compiled = vybe_compiler::dynamic::RuntimeCompilerService::new(&mut vm)
        .compile_path(&path).unwrap();
    assert_no_dynamic_equality(&compiled.chunks);
    let bytes = vybe_platform_wasm::write_wasm_with_options(&compiled.chunks,
        vybe_platform_wasm::WasmWriteOptions {
            include_vybe_metadata: false,
            include_name_section: true,
        });
    validate_wasm_with_context(&bytes, &compiled.chunks);
    vybe_compiler::dynamic::RuntimeCompilerService::new(&mut vm)
        .run_compiled(compiled).unwrap();
    assert_eq!(vm.pending_exit_code, 0);
}

#[test]
fn record_array_memset_preserves_backing_and_evaluates_owner_once() {
    let path = Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../tests/c/doom_startup/record_array_memset.c");
    vybe_language_c::register();
    let mut vm = vybe_runtime::VM::new();
    vybe_runtime::init_all_registered(&mut vm, &vybe_runtime::capabilities::Capabilities::all());
    let compiled = vybe_compiler::dynamic::RuntimeCompilerService::new(&mut vm)
        .compile_path(&path).unwrap();
    assert_no_dynamic_equality(&compiled.chunks);
    let bytes = vybe_platform_wasm::write_wasm_with_options(&compiled.chunks,
        vybe_platform_wasm::WasmWriteOptions {
            include_vybe_metadata: false,
            include_name_section: true,
        });
    validate_wasm_with_context(&bytes, &compiled.chunks);
    vybe_compiler::dynamic::RuntimeCompilerService::new(&mut vm)
        .run_compiled(compiled).unwrap();
    assert_eq!(vm.pending_exit_code, 0);
}

#[test]
fn header_scalar_and_signal_guards_do_not_use_dynamic_equality() {
    for fixture in ["sdl_modifier_comparison", "signal_handler_guard"] {
        let path = Path::new(env!("CARGO_MANIFEST_DIR"))
            .join(format!("../../tests/c/doom_startup/{fixture}.c"));
        vybe_language_c::register();
        let mut vm = vybe_runtime::VM::new();
        vybe_runtime::init_all_registered(&mut vm, &vybe_runtime::capabilities::Capabilities::all());
        let compiled = vybe_compiler::dynamic::RuntimeCompilerService::new(&mut vm)
            .compile_path(&path).unwrap();
        assert_no_dynamic_equality(&compiled.chunks);
        let bytes = vybe_platform_wasm::write_wasm_with_options(&compiled.chunks,
            vybe_platform_wasm::WasmWriteOptions {
                include_vybe_metadata: false,
                include_name_section: true,
            });
        validate_wasm_with_context(&bytes, &compiled.chunks);
        vybe_compiler::dynamic::RuntimeCompilerService::new(&mut vm)
            .run_compiled(compiled).unwrap();
        assert_eq!(vm.pending_exit_code, 0);
    }
}

#[test]
fn doom_projection_preserves_unsigned_angles_and_wall_scales() {
    for fixture in ["unsigned_angle_wrap", "doom_projection_math"] {
        let path = Path::new(env!("CARGO_MANIFEST_DIR"))
            .join(format!("../../tests/c/doom_startup/{fixture}.c"));
        vybe_language_c::register();
        let mut vm = vybe_runtime::VM::new();
        vybe_runtime::init_all_registered(&mut vm, &vybe_runtime::capabilities::Capabilities::all());
        let compiled = vybe_compiler::dynamic::RuntimeCompilerService::new(&mut vm)
            .compile_path(&path).unwrap();
        assert_no_dynamic_equality(&compiled.chunks);
        let bytes = vybe_platform_wasm::write_wasm_with_options(&compiled.chunks,
            vybe_platform_wasm::WasmWriteOptions {
                include_vybe_metadata: false,
                include_name_section: true,
            });
        validate_wasm_with_context(&bytes, &compiled.chunks);
        vybe_compiler::dynamic::RuntimeCompilerService::new(&mut vm)
            .run_compiled(compiled).unwrap();
        assert_eq!(vm.pending_exit_code, 0);
    }
}

#[test]
fn floating_array_pointer_reads_keep_static_comparisons_and_places() {
    let path = Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../tests/c/doom_startup/floating_array_pointer_reads.c");
    vybe_language_c::register();
    let mut vm = vybe_runtime::VM::new();
    vybe_runtime::init_all_registered(&mut vm, &vybe_runtime::capabilities::Capabilities::all());
    let compiled = vybe_compiler::dynamic::RuntimeCompilerService::new(&mut vm)
        .compile_path(&path).unwrap();
    assert_no_dynamic_equality(&compiled.chunks);
    let bytes = vybe_platform_wasm::write_wasm_with_options(&compiled.chunks,
        vybe_platform_wasm::WasmWriteOptions {
            include_vybe_metadata: false,
            include_name_section: true,
        });
    validate_wasm_with_context(&bytes, &compiled.chunks);
    vybe_compiler::dynamic::RuntimeCompilerService::new(&mut vm)
        .run_compiled(compiled).unwrap();
    assert_eq!(vm.pending_exit_code, 0);
}

#[test]
fn pointer_strides_preserve_local_and_linear_backing_semantics() {
    let path = Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../tests/c/doom_startup/pointer_stride_matrix.c");
    vybe_language_c::register();
    let mut vm = vybe_runtime::VM::new();
    vybe_runtime::init_all_registered(&mut vm, &vybe_runtime::capabilities::Capabilities::all());
    let compiled = vybe_compiler::dynamic::RuntimeCompilerService::new(&mut vm)
        .compile_path(&path).unwrap();
    assert_no_dynamic_equality(&compiled.chunks);
    let bytes = vybe_platform_wasm::write_wasm_with_options(&compiled.chunks,
        vybe_platform_wasm::WasmWriteOptions {
            include_vybe_metadata: false,
            include_name_section: true,
        });
    validate_wasm_with_context(&bytes, &compiled.chunks);
    vybe_compiler::dynamic::RuntimeCompilerService::new(&mut vm)
        .run_compiled(compiled).unwrap();
    assert_eq!(vm.pending_exit_code, 0);
}

#[test]
fn doom_hud_nested_lines_preserve_dirty_counter_and_border_erases() {
    for fixture in ["record_array_member_updates", "doom_hud_erase"] {
        let path = Path::new(env!("CARGO_MANIFEST_DIR"))
            .join(format!("../../tests/c/doom_startup/{fixture}.c"));
        vybe_language_c::register();
        let mut vm = vybe_runtime::VM::new();
        vybe_runtime::init_all_registered(&mut vm, &vybe_runtime::capabilities::Capabilities::all());
        let compiled = vybe_compiler::dynamic::RuntimeCompilerService::new(&mut vm)
            .compile_path(&path).unwrap();
        assert_no_dynamic_equality(&compiled.chunks);
        let bytes = vybe_platform_wasm::write_wasm_with_options(&compiled.chunks,
            vybe_platform_wasm::WasmWriteOptions {
                include_vybe_metadata: false,
                include_name_section: true,
            });
        validate_wasm_with_context(&bytes, &compiled.chunks);
        vybe_compiler::dynamic::RuntimeCompilerService::new(&mut vm)
            .run_compiled(compiled).unwrap();
        assert_eq!(vm.pending_exit_code, 0);
    }
}

#[test]
fn doom_draw_kernels_preserve_near_wall_mapping_and_guards() {
    for fixture in ["doom_draw_kernel_edges", "doom_draw_kernel_heap"] {
        let path = Path::new(env!("CARGO_MANIFEST_DIR"))
            .join(format!("../../tests/c/doom_startup/{fixture}.c"));
        assert_native_map_fixture(&path);
        vybe_language_c::register();
        let mut vm = vybe_runtime::VM::new();
        vybe_runtime::init_all_registered(&mut vm, &vybe_runtime::capabilities::Capabilities::all());
        let compiled = vybe_compiler::dynamic::RuntimeCompilerService::new(&mut vm)
            .compile_path(&path).unwrap();
        assert_no_dynamic_equality(&compiled.chunks);
        let bytes = vybe_platform_wasm::write_wasm_with_options(&compiled.chunks,
            vybe_platform_wasm::WasmWriteOptions {
                include_vybe_metadata: false,
                include_name_section: true,
            });
        validate_wasm_with_context(&bytes, &compiled.chunks);
        vybe_compiler::dynamic::RuntimeCompilerService::new(&mut vm)
            .run_compiled(compiled).unwrap();
        assert_eq!(vm.pending_exit_code, 0);
    }
}

#[test]
fn pickup_blocklist_traversal_preserves_links_during_removal() {
    let root = Path::new(env!("CARGO_MANIFEST_DIR")).join("../..");
    let source = root.join("tests/c/doom_startup/doom_pickup_blocklist.c");
    assert_native_map_fixture(&source);
    vybe_language_c::register();
    let mut vm = vybe_runtime::VM::new();
    vybe_runtime::init_all_registered(&mut vm, &vybe_runtime::capabilities::Capabilities::all());
    let compiled = vybe_compiler::dynamic::RuntimeCompilerService::new(&mut vm)
        .compile_path(&source).unwrap();
    assert_no_dynamic_equality(&compiled.chunks);
    let bytes = vybe_platform_wasm::write_wasm_with_options(&compiled.chunks,
        vybe_platform_wasm::WasmWriteOptions {
            include_vybe_metadata: false,
            include_name_section: true,
        });
    validate_wasm_with_context(&bytes, &compiled.chunks);
    vybe_compiler::dynamic::RuntimeCompilerService::new(&mut vm)
        .run_compiled(compiled).unwrap();
    assert_eq!(vm.pending_exit_code, 0);
}

fn assert_native_map_fixture(source: &Path) {
    assert_native_map_fixture_with_flags(source, &[]);
}

fn assert_native_map_fixture_with_flags(source: &Path, flags: &[&str]) {
    let root = Path::new(env!("CARGO_MANIFEST_DIR")).join("../..");
    struct NativeArtifact(std::path::PathBuf);
    impl Drop for NativeArtifact {
        fn drop(&mut self) { let _ = std::fs::remove_file(&self.0); }
    }
    let reference = NativeArtifact(std::env::temp_dir().join(format!(
        "vybe-map-reference-{}-{}", std::process::id(),
        source.file_stem().unwrap().to_string_lossy())));
    let compiler = std::env::var_os("CC").unwrap_or_else(|| "cc".into());
    let output = std::process::Command::new(compiler)
        .args(["-std=gnu11", "-O0", "-ffunction-sections", "-fdata-sections"])
        .args(flags)
        .arg("-idirafter").arg(root.join("tools/compat"))
        .arg(format!("-I{}", root.join("examples/chocolate-doom/src").display()))
        .arg(format!("-I{}", root.join("examples/chocolate-doom/src/doom").display()))
        .arg(&source)
        .arg(if cfg!(target_os = "macos") { "-Wl,-dead_strip" } else { "-Wl,--gc-sections" })
        .arg("-o").arg(&reference.0).output().unwrap();
    assert!(output.status.success(), "native compilation: {}", String::from_utf8_lossy(&output.stderr));
    let output = std::process::Command::new(&reference.0).output().unwrap();
    assert!(output.status.success(), "native map assertions for {}: {}\n{}", source.display(),
        String::from_utf8_lossy(&output.stdout), String::from_utf8_lossy(&output.stderr));
}

#[test]
fn momentum_wall_sliding_matches_native_position_and_links() {
    let source = Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../tests/c/doom_startup/doom_momentum_wall.c");
    assert_native_map_fixture(&source);
    vybe_language_c::register();
    let mut vm = vybe_runtime::VM::new();
    vybe_runtime::init_all_registered(&mut vm, &vybe_runtime::capabilities::Capabilities::all());
    let compiled = vybe_compiler::dynamic::RuntimeCompilerService::new(&mut vm)
        .compile_path(&source).unwrap();
    assert_no_dynamic_equality(&compiled.chunks);
    let bytes = vybe_platform_wasm::write_wasm_with_options(&compiled.chunks,
        vybe_platform_wasm::WasmWriteOptions {
            include_vybe_metadata: false,
            include_name_section: true,
        });
    validate_wasm_with_context(&bytes, &compiled.chunks);
    vybe_compiler::dynamic::RuntimeCompilerService::new(&mut vm)
        .run_compiled(compiled).unwrap();
    assert_eq!(vm.pending_exit_code, 0);
}

fn check_transition_fixture(fixture: &str) {
    check_transition_fixture_with_flags(fixture, &[]);
}

fn check_transition_fixture_with_flags(fixture: &str, flags: &[&str]) {
    let source = Path::new(env!("CARGO_MANIFEST_DIR"))
        .join(format!("../../tests/c/doom_startup/{fixture}.c"));
    assert_native_map_fixture_with_flags(&source, flags);
    vybe_language_c::register();
    let mut vm = vybe_runtime::VM::new();
    vybe_runtime::init_all_registered(&mut vm, &vybe_runtime::capabilities::Capabilities::all());
    let compiled = vybe_compiler::dynamic::RuntimeCompilerService::new(&mut vm)
        .compile_path(&source).unwrap();
    assert_no_dynamic_equality(&compiled.chunks);
    let bytes = vybe_platform_wasm::write_wasm_with_options(&compiled.chunks,
        vybe_platform_wasm::WasmWriteOptions {
            include_vybe_metadata: false,
            include_name_section: true,
        });
    validate_wasm_with_context(&bytes, &compiled.chunks);
    vybe_compiler::dynamic::RuntimeCompilerService::new(&mut vm)
        .run_compiled(compiled).unwrap();
    assert_eq!(vm.pending_exit_code, 0, "{fixture}: C assertions");
}

#[test]
fn door_plane_transitions_match_native() {
    check_transition_fixture("doom_door_plane_transitions");
}

#[test]
fn door_use_trace_matches_native() {
    check_transition_fixture("doom_use_door_trace");
}

#[test]
fn bounded_string_formatting_preserves_byte_offsets_and_single_evaluation() {
    check_transition_fixture("printf_string_precision_bytes");
}

#[test]
fn bounded_string_precision_never_reads_beyond_linear_memory() {
    check_transition_fixture_with_flags("printf_precision_memory_end",
        &["-DNATIVE_GUARD_PAGE", "-D_DARWIN_C_SOURCE", "-D_DEFAULT_SOURCE"]);
}

#[test]
fn doom_file_base_extraction_preserves_caller_storage_and_guards() {
    check_transition_fixture("doom_extract_file_base_precision");
}

#[test]
fn string_backed_pointer_views_preserve_lengths_offsets_and_formatting() {
    check_transition_fixture("string_backed_pointer_offsets");
}

#[test]
fn byte_lookup_views_preserve_live_storage_and_descriptor_mutations() {
    check_transition_fixture("byte_lookup_view_stability");
}

#[test]
fn sector_height_propagation_matches_native() {
    check_transition_fixture("doom_sector_height_change");
}

#[test]
fn pickup_effects_and_deferred_removal_match_native() {
    check_transition_fixture("doom_pickup_effects");
}

#[test]
fn pickup_blocklink_insert_matches_native() {
    check_transition_fixture("doom_pickup_blocklink_insert");
}

#[test]
fn zero_tic_actor_state_transitions_match_native() {
    check_transition_fixture("doom_state_transitions");
}

#[test]
fn doom_keyboard_release_updates_gameplay_commands_match_native() {
    check_transition_fixture("doom_keyboard_command_release");
}

#[test]
fn doom_event_ring_boundaries_match_native_queue_hazard() {
    check_transition_fixture("doom_event_ring_boundaries");
}

#[test]
fn enemy_player_search_termination_matches_native() {
    check_transition_fixture("doom_enemy_player_search");
}

#[test]
fn platform_removal_preserves_sector_and_thinker_identity() {
    check_transition_fixture("doom_platform_removal");
}

#[test]
fn embedded_record_addresses_preserve_identity_layout_and_evaluation() {
    check_transition_fixture("embedded_record_address_identity");
}

#[test]
fn pointer_field_rhs_evaluates_once_and_returns_assigned_pointer() {
    check_transition_fixture("record_pointer_field_rhs_once");
}

#[test]
fn repeated_blocklist_moves_preserve_links_and_terminate() {
    check_transition_fixture("doom_blocklist_moves");
}

#[test]
fn indexed_record_addresses_preserve_singleton_and_array_aliases() {
    for fixture in ["single_record_index_address", "doom_blockmap_wall", "doom_trymove_wall"] {
        let path = Path::new(env!("CARGO_MANIFEST_DIR"))
            .join(format!("../../tests/c/doom_startup/{fixture}.c"));
        assert_native_map_fixture(&path);
        vybe_language_c::register();
        let mut vm = vybe_runtime::VM::new();
        vybe_runtime::init_all_registered(&mut vm, &vybe_runtime::capabilities::Capabilities::all());
        let compiled = vybe_compiler::dynamic::RuntimeCompilerService::new(&mut vm)
            .compile_path(&path).unwrap();
        assert_no_dynamic_equality(&compiled.chunks);
        let bytes = vybe_platform_wasm::write_wasm_with_options(&compiled.chunks,
            vybe_platform_wasm::WasmWriteOptions {
                include_vybe_metadata: false,
                include_name_section: true,
            });
        validate_wasm_with_context(&bytes, &compiled.chunks);
        vybe_compiler::dynamic::RuntimeCompilerService::new(&mut vm)
            .run_compiled(compiled).unwrap();
        assert_eq!(vm.pending_exit_code, 0);
    }
}

fn canvas_import_signatures(bytes: &[u8]) -> Vec<(u32, String, usize, usize)> {
        let mut canvas = Vec::new();
        let mut signatures = Vec::new();
        let mut import_index = 0;
        for payload in wasmparser::Parser::new(0).parse_all(bytes) {
            match payload.unwrap() {
                wasmparser::Payload::TypeSection(section) => {
                    for group in section {
                        for ty in group.unwrap().into_types() {
                            signatures.push(match ty.composite_type.inner {
                                wasmparser::CompositeInnerType::Func(f) =>
                                    Some((f.params().len(), f.results().len())),
                                _ => None,
                            });
                        }
                    }
                }
                wasmparser::Payload::ImportSection(section) => {
                    for import in section.into_imports() {
                        let import = import.unwrap();
                        if let wasmparser::TypeRef::Func(ty) | wasmparser::TypeRef::FuncExact(ty) = import.ty {
                            if import.module == "web:canvas" {
                                let (params, results) = signatures[ty as usize].unwrap();
                                canvas.push((import_index, import.name.to_string(), params, results));
                            }
                            import_index += 1;
                        }
                    }
                }
                _ => {},
            }
        }
        canvas
}

fn validate_wasm_with_context(bytes: &[u8], chunks: &[vybe_runtime::Chunk]) {
    if let Err(error) = wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
        .validate_all(bytes)
    {
        for (index, name, params, results) in canvas_import_signatures(bytes) {
            eprintln!("canvas import {index}: {name} ({params}, {results})");
        }
        let mut body_index = 0;
        for payload in wasmparser::Parser::new(0).parse_all(bytes) {
            if let wasmparser::Payload::CodeSectionEntry(body) = payload.unwrap() {
                if body.range().contains(&error.offset()) {
                    let mut reader = body.get_operators_reader().unwrap();
                    let mut nearby = Vec::new();
                    while !reader.eof() {
                        let offset = reader.original_position();
                        let operation = reader.read().unwrap();
                        if offset + 40 >= error.offset() && offset <= error.offset() + 8 {
                            nearby.push(format!("{offset}: {operation:?}"));
                        }
                    }
                    let chunk = chunks.get(body_index);
                    panic!("{error}; body {body_index}, chunk {:?}, signature {:?}, i64 locals {:?}, names {:?}: {}",
                        chunk.map(|chunk| &chunk.name),
                        chunk.map(|chunk| &chunk.func_sig),
                        chunk.map(|chunk| &chunk.i64_locals),
                        chunk.map(|chunk| &chunk.local_names), nearby.join("\n"));
                }
                body_index += 1;
            }
        }
        panic!("{error}");
    }
    let mut abstract_imports = Vec::new();
    for payload in wasmparser::Parser::new(0).parse_all(bytes) {
        match payload.unwrap() {
            wasmparser::Payload::ImportSection(section) => {
                for import in section.into_imports() {
                    let import = import.unwrap();
                    if matches!(import.ty,
                        wasmparser::TypeRef::Func(_) | wasmparser::TypeRef::FuncExact(_))
                    {
                        abstract_imports.push(import.module == "ecma:value"
                            && matches!(import.name, "abstractEq" | "abstractNe"));
                    }
                }
            }
            wasmparser::Payload::CodeSectionEntry(body) => {
                let mut reader = body.get_operators_reader().unwrap();
                while !reader.eof() {
                    let offset = reader.original_position();
                    if let wasmparser::Operator::Call { function_index }
                        | wasmparser::Operator::ReturnCall { function_index }
                        | wasmparser::Operator::RefFunc { function_index } = reader.read().unwrap()
                    {
                        assert_ne!(abstract_imports.get(function_index as usize), Some(&true),
                            "emitted WASM references abstract equality at byte {offset}, function {function_index}");
                    }
                }
            }
            wasmparser::Payload::CustomSection(section) => {
                if let wasmparser::KnownCustom::Name(reader) = section.as_known() {
                    for subsection in reader {
                        subsection.expect("emitted debug name subsections must be well formed");
                    }
                }
            }
            _ => {}
        }
    }
}

#[test]
fn computed_record_store_uses_emitted_typedef_type() {
    let module = vybe_language_c::parse(
        "typedef struct LinkTag { struct LinkTag *next; } Link;\n\
         Link head; void attach(Link *node) { head.next->next = node; }\n\
         int main(void) { return 0; }",
    ).unwrap();
    let body = module.body.iter().find_map(|statement| match &statement.kind {
        vybe_ast::StmtKind::FunctionDecl { name, body, .. } if name == "attach" => Some(body),
        _ => None,
    }).unwrap();
    let mut tag_stores = 0;
    let mut declared_stores = 0;
    for statement in body {
        statement.walk_exprs(&mut |expression| {
            if let vybe_ast::ExprKind::Assign { target, .. } = &expression.kind
                && let vybe_ast::ExprKind::Member { object, field, .. } = &target.kind
                && field == "next"
                && let vybe_ast::ExprKind::Cast { type_name, .. } = &object.kind {
                tag_stores += usize::from(type_name == "LinkTag");
                declared_stores += usize::from(type_name == "Link");
            }
        });
    }
    assert_eq!(tag_stores, 0, "store names a tag with no emitted record class");
    assert!(declared_stores > 0, "store must retain the emitted concrete type");
}

#[test]
fn function_scope_global_record_address_keeps_its_place() {
    let module = vybe_language_c::parse(
        "struct link { struct link *next; }; struct link head;\n\
         void initialize(void) { head.next = &head; }\n\
         int main(void) { initialize(); return 0; }",
    ).unwrap();
    let body = module.body.iter().find_map(|statement| match &statement.kind {
        vybe_ast::StmtKind::FunctionDecl { name, body, .. } if name == "initialize" => Some(body),
        _ => None,
    }).unwrap();
    let mut keeps_place = false;
    for statement in body {
        statement.walk_exprs(&mut |expression| {
            if let vybe_ast::ExprKind::Unary { op: vybe_ast::UnaryOp::AddrOf, expr } = &expression.kind {
                keeps_place |= matches!(&expr.kind, vybe_ast::ExprKind::Ident(name) if name == "head");
            }
        });
    }
    assert!(keeps_place, "global record address was lowered to a value-backed wrapper");
}

#[test]
fn threaded_thinker_head_keeps_managed_storage() {
    let path = Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../tests/c/doom_startup/threaded_thinker_head.c");
    vybe_language_c::register();
    let mut vm = vybe_runtime::VM::new();
    vybe_runtime::init_all_registered(&mut vm, &vybe_runtime::capabilities::Capabilities::all());
    let compiled = vybe_compiler::dynamic::RuntimeCompilerService::new(&mut vm)
        .compile_path(&path).unwrap();
    vybe_compiler::dynamic::RuntimeCompilerService::new(&mut vm)
        .run_compiled(compiled).unwrap();
    assert_eq!(vm.pending_exit_code, 0);
}

#[test]
fn sdl_boolean_constants_use_integer_values() {
    vybe_language_c::register();
    let path = Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../tests/c/doom_startup/sdl_boolean_constants.c");
    let mut vm = vybe_runtime::VM::new();
    vybe_runtime::init_all_registered(&mut vm, &vybe_runtime::capabilities::Capabilities::all());
    let compiled = vybe_compiler::dynamic::RuntimeCompilerService::new(&mut vm)
        .compile_path(&path).unwrap();
    let main = compiled.chunks.iter().find(|chunk| chunk.name == "main").unwrap();
    assert!(!vybe_runtime::debug::disassemble(main).contains("dyneq"));
    let bytes = vybe_platform_wasm::write_wasm_with_options(&compiled.chunks,
        vybe_platform_wasm::WasmWriteOptions {
            include_vybe_metadata: false,
            include_name_section: true,
        });
    wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
        .validate_all(&bytes).unwrap();
    vybe_compiler::dynamic::RuntimeCompilerService::new(&mut vm)
        .run_compiled(compiled).unwrap();
    assert_eq!(vm.pending_exit_code, 0);
}

#[test]
fn mixed_pointer_result_field_transfer_preserves_linear_address() {
    vybe_language_c::register();
    let path = Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../tests/c/doom_startup/mixed_pointer_result_field_transfer.c");
    let chunks = vybe_compiler::projects::load(&path).unwrap().compile().unwrap();
    let bytes = vybe_platform_wasm::write_wasm_with_options(&chunks,
        vybe_platform_wasm::WasmWriteOptions {
            include_vybe_metadata: false,
            include_name_section: true,
        });
    wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
        .validate_all(&bytes).unwrap();
    let mut vm = vybe_runtime::VM::new();
    vybe_runtime::init_all_registered(&mut vm, &vybe_runtime::capabilities::Capabilities::all());
    vm.run(chunks).unwrap();
    assert_eq!(vm.pending_exit_code, 0);
}

#[test]
fn large_static_zero_initializer_exports_bounded_string_constants() {
    vybe_language_c::register();
    let path = Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../tests/c/doom_startup/large_static_zero_initializer.c");
    let chunks = vybe_compiler::projects::load(&path).unwrap().compile().unwrap();
    let bytes = vybe_platform_wasm::write_wasm(&chunks);
    wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
        .validate_all(&bytes).unwrap();
}

#[test]
fn character_array_initializers_and_copies_preserve_byte_storage() {
    vybe_language_c::register();
    for name in ["char_array_numeric_initializer_copy.c", "cheat_sequence.c",
        "record_char_array_initializer.c", "texture_hash_table_high_bucket.c",
        "unsigned_division_return.c"] {
        let path = Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("../../tests/c/doom_startup").join(name);
        eprintln!("byte-storage fixture: {name}");
        let chunks = vybe_compiler::projects::load(&path).unwrap().compile().unwrap();
        if name == "unsigned_division_return.c" {
            for (function, opcode) in [("word_div", "i32.div_u"), ("word_rem", "i32.rem_u")] {
                let chunk = chunks.iter().find(|chunk| chunk.name == function).unwrap();
                let dump = vybe_runtime::debug::disassemble(chunk);
                assert!(dump.contains(opcode), "{function}: {dump}");
                assert!(!dump.contains("f64.div"), "{function}: {dump}");
            }
        }
        let bytes = vybe_platform_wasm::write_wasm_with_options(&chunks,
            vybe_platform_wasm::WasmWriteOptions {
                include_vybe_metadata: false,
                include_name_section: true,
            });
        wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
            .validate_all(&bytes).unwrap_or_else(|error| panic!("{name}: {error}"));
        let mut vm = vybe_runtime::VM::new();
        vybe_runtime::init_all_registered(&mut vm, &vybe_runtime::capabilities::Capabilities::all());
        vm.run(chunks.clone()).unwrap_or_else(|error| panic!("{name}: {error}"));
        assert_eq!(vm.pending_exit_code, 0, "{name}");
        assert_no_dynamic_equality(&chunks);
    }
}

#[test]
fn sprite_frame_addresses_use_record_byte_stride() {
    let path = Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../tests/c/doom_startup/sprite_frame_table.c");
    vybe_language_c::register();
    let chunks = vybe_compiler::projects::load(&path).unwrap().compile().unwrap();
    let bytes = vybe_platform_wasm::write_wasm_with_options(&chunks,
        vybe_platform_wasm::WasmWriteOptions {
            include_vybe_metadata: false,
            include_name_section: true,
        });
    wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
        .validate_all(&bytes).unwrap();
    let mut vm = vybe_runtime::VM::new();
    vybe_runtime::init_all_registered(&mut vm, &vybe_runtime::capabilities::Capabilities::all());
    vm.run(chunks).unwrap();
    assert_eq!(vm.pending_exit_code, 0);
}

#[test]
fn allocator_first_member_preserves_callback_table() {
    let path = Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../tests/c/doom_startup/struct_pointer_from_void_allocator_nested_function_table.c");
    vybe_language_c::register();
    let chunks = vybe_compiler::projects::load(&path).unwrap().compile().unwrap();
    let bytes = vybe_platform_wasm::write_wasm_with_options(&chunks,
        vybe_platform_wasm::WasmWriteOptions {
            include_vybe_metadata: false,
            include_name_section: true,
        });
    wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
        .validate_all(&bytes).unwrap();
    let mut vm = vybe_runtime::VM::new();
    vybe_runtime::init_all_registered(&mut vm, &vybe_runtime::capabilities::Capabilities::all());
    vm.run(chunks).unwrap();
    assert_eq!(vm.pending_exit_code, 0);
}

#[test]
fn file_registry_helpers_use_static_comparisons() {
    let path = Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../tests/c/doom_startup/static_file_registry.c");
    vybe_language_c::register();
    let chunks = vybe_compiler::projects::load(&path).unwrap().compile().unwrap();
    assert_no_dynamic_equality(&chunks);
    for name in [
        "__c_fs_is_file", "__c_path_present", "__c_remove_h", "__c_rename_h",
        "__c_fopen_h", "__c_fsync_h", "__c_ungetc_h", "__c_ftell_h", "__c_srand_h",
        "__c_fwrite_h",
    ] {
        let chunk = chunks.iter().find(|chunk| chunk.name == name)
            .unwrap_or_else(|| panic!("fixture must exercise {name}"));
        let dump = vybe_runtime::debug::disassemble(chunk);
        assert!(!dump.contains("dyneq"), "dynamic equality remains in {name}: {dump}");
    }
    for chunk in chunks.iter().filter(|chunk| chunk.name == "<lambda>") {
        let dump = vybe_runtime::debug::disassemble(chunk);
        assert!(!dump.contains("dyneq"), "dynamic equality in generated file adapter: {dump}");
    }
    let bytes = vybe_platform_wasm::write_wasm_with_options(&chunks,
        vybe_platform_wasm::WasmWriteOptions {
            include_vybe_metadata: false,
            include_name_section: true,
        });
    validate_wasm_with_context(&bytes, &chunks);
    let mut vm = vybe_runtime::VM::new();
    vybe_runtime::init_all_registered(&mut vm, &vybe_runtime::capabilities::Capabilities::all());
    vm.run(chunks).unwrap();
    assert_eq!(vm.pending_exit_code, 0);
}

#[test]
fn variadic_record_pointer_conditions_validate() {
    let path = Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../tests/c/doom_startup/multi_glob_conditions.c");
    vybe_language_c::register();
    let chunks = vybe_compiler::projects::load(&path).unwrap().compile().unwrap();
    assert_no_dynamic_equality(&chunks);
    let bytes = vybe_platform_wasm::write_wasm_with_options(&chunks,
        vybe_platform_wasm::WasmWriteOptions {
            include_vybe_metadata: false,
            include_name_section: true,
        });
    validate_wasm_with_context(&bytes, &chunks);
    let mut vm = vybe_runtime::VM::new();
    vybe_runtime::init_all_registered(&mut vm, &vybe_runtime::capabilities::Capabilities::all());
    let result = vm.run(chunks);
    let stage = vm.globals_by_name().into_iter()
        .find(|(name, _)| name == "glob_stage");
    assert!(result.is_ok(), "glob stage {stage:?}: {result:?}");
    assert_eq!(vm.pending_exit_code, 0);
}

#[test]
fn sector_line_tables_survive_zone_level_restarts() {
    vybe_language_c::register();
    for fixture in ["sector_line_tables.c", "zone_sector_line_tables.c"] {
        let path = Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("../../tests/c/doom_startup").join(fixture);
        let chunks = vybe_compiler::projects::load(&path).unwrap().compile().unwrap();
        assert_no_dynamic_equality(&chunks);
        let bytes = vybe_platform_wasm::write_wasm_with_options(&chunks,
            vybe_platform_wasm::WasmWriteOptions {
                include_vybe_metadata: false,
                include_name_section: true,
            });
        validate_wasm_with_context(&bytes, &chunks);
        let mut vm = vybe_runtime::VM::new();
        vybe_runtime::init_all_registered(&mut vm, &vybe_runtime::capabilities::Capabilities::all());
        let table_dump = chunks.iter().find(|chunk| chunk.name == "group_lines")
            .map(vybe_runtime::debug::disassemble).unwrap_or_default();
        vm.run(chunks).unwrap_or_else(|error| panic!("{fixture}: {error:?}\n{table_dump}"));
        assert_eq!(vm.pending_exit_code, 0, "{fixture}");
    }
}

#[test]
fn doom_cheat_size_fields_preserve_record_backing() {
    vybe_language_c::register();
    let path = Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../tests/c/doom_startup/doom_cheat_sequence.c");
    let chunks = vybe_compiler::projects::load(&path).unwrap().compile().unwrap();
    assert_no_dynamic_equality(&chunks);
    let bytes = vybe_platform_wasm::write_wasm_with_options(&chunks,
        vybe_platform_wasm::WasmWriteOptions {
            include_vybe_metadata: false,
            include_name_section: true,
        });
    validate_wasm_with_context(&bytes, &chunks);
    let mut vm = vybe_runtime::VM::new();
    vybe_runtime::init_all_registered(&mut vm, &vybe_runtime::capabilities::Capabilities::all());
    vm.run(chunks).unwrap();
    assert_eq!(vm.pending_exit_code, 0);
}

#[test]
fn nested_widget_addresses_preserve_parent_storage() {
    let path = Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../tests/c/doom_startup/nested_widget_address.c");
    vybe_language_c::register();
    let chunks = vybe_compiler::projects::load(&path).unwrap().compile().unwrap();
    for chunk in chunks.iter().filter(|chunk| chunk.name == "main" || chunk.name == "<lambda>") {
        let dump = vybe_runtime::debug::disassemble(chunk);
        assert!(!dump.contains("dyneq"), "dynamic equality in nested pointer comparison: {dump}");
    }
    let mut vm = vybe_runtime::VM::new();
    vybe_runtime::init_all_registered(&mut vm, &vybe_runtime::capabilities::Capabilities::all());
    vm.run(chunks).unwrap();
    assert_eq!(vm.pending_exit_code, 0);
}

#[test]
fn sdl_texture_output_stores_export_with_balanced_stack() {
    vybe_language_c::register();
    let path = Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../tests/c/doom_startup/sdl_output_stack.c");
    let chunks = vybe_compiler::projects::load(&path).unwrap().compile().unwrap();
    assert_no_dynamic_equality(&chunks);
    let bytes = vybe_platform_wasm::write_wasm_with_options(&chunks,
        vybe_platform_wasm::WasmWriteOptions {
            include_vybe_metadata: false,
            include_name_section: true,
        });
    validate_wasm_with_context(&bytes, &chunks);
    let overloads: std::collections::BTreeSet<_> = canvas_import_signatures(&bytes)
        .into_iter().filter(|(_, name, _, _)| name == "drawImage")
        .map(|(_, _, params, results)| (params, results)).collect();
    assert_eq!(overloads, [(4, 0), (10, 0)].into_iter().collect());
}

#[test]
fn canvas_overloads_preserve_root_imports_in_leaf_chunks() {
    use vybe_runtime::{Chunk, Op};
    let mut root = Chunk::new("<canvas-overloads>");
    let draw = root.add_import("web:canvas", "drawImage");
    let put = root.add_import("web:canvas", "putImageData");
    root.emit_i32_const(0, 0);
    root.emit_op(Op::RETURN, 0);
    let mut leaf = Chunk::new("canvas-overloads");
    for (import, arity) in [(draw, 4), (draw, 6), (draw, 10), (put, 4), (put, 8)] {
        for _ in 0..arity {
            leaf.emit_i32_const(0, 0);
        }
        leaf.emit_call(import, arity, 0);
        leaf.emit_op(Op::DROP, 0);
    }
    leaf.emit_i32_const(0, 0);
    leaf.emit_op(Op::RETURN, 0);
    let chunks = vec![root, leaf];
    let bytes = vybe_platform_wasm::write_wasm_with_options(&chunks,
        vybe_platform_wasm::WasmWriteOptions {
            include_vybe_metadata: false,
            include_name_section: true,
        });
    validate_wasm_with_context(&bytes, &chunks);
    let signatures = canvas_import_signatures(&bytes);
    for (name, expected) in [
        ("drawImage", vec![(4, 0), (6, 0), (10, 0)]),
        ("putImageData", vec![(4, 0), (8, 0)]),
    ] {
        let actual: std::collections::BTreeSet<_> = signatures.iter()
            .filter(|(_, method, _, _)| method == name)
            .map(|(_, _, params, results)| (*params, *results)).collect();
        assert_eq!(actual, expected.into_iter().collect(), "{name}");
    }
}

#[test]
fn linear_pointer_slots_preserve_indexed_backing() {
    let path = Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../tests/c/doom_startup/linear_pointer_slot_index.c");
    vybe_language_c::register();
    let chunks = vybe_compiler::projects::load(&path).unwrap().compile().unwrap();
    let bytes = vybe_platform_wasm::write_wasm_with_options(&chunks,
        vybe_platform_wasm::WasmWriteOptions {
            include_vybe_metadata: false,
            include_name_section: true,
        });
    wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
        .validate_all(&bytes).unwrap();
    let main = chunks.iter().find(|chunk| chunk.name == "main").unwrap();
    let dump = vybe_runtime::debug::disassemble(main);
    assert!(!dump.contains("dyneq"), "dynamic pointer comparison: {dump}");
    let mut vm = vybe_runtime::VM::new();
    vybe_runtime::init_all_registered(&mut vm, &vybe_runtime::capabilities::Capabilities::all());
    vm.run(chunks).unwrap_or_else(|error| panic!("{error}\n{dump}"));
    assert_eq!(vm.pending_exit_code, 0);
}

#[test]
fn callback_parameter_prototypes_preserve_functions() {
    let path = Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../tests/c/doom_startup/callback_parameter_prototype.c");
    vybe_language_c::register();
    let chunks = vybe_compiler::projects::load(&path).unwrap().compile().unwrap();
    let mut vm = vybe_runtime::VM::new();
    vybe_runtime::init_all_registered(&mut vm, &vybe_runtime::capabilities::Capabilities::all());
    vm.run(chunks).unwrap();
    assert_eq!(vm.pending_exit_code, 0);
}

#[test]
fn linear_character_search_emits_valid_wasm() {
    use vybe_runtime::{Chunk, Op};
    for (needle, expected) in [(58, 33), (0, 37), (122, 0), (-1, 32)] {
        let mut chunk = Chunk::new("<script>");
        chunk.memory_min_pages = vec![1];
        for (offset, byte) in [255, 58, 98, 58, 99, 0].into_iter().enumerate() {
            chunk.emit_i32_const(32 + offset as i32, 0);
            chunk.emit_i32_const(byte, 0);
            chunk.emit_op(Op::I32_STORE8, 0);
        }
        chunk.emit_i32_const(32, 0);
        chunk.emit_i32_const(needle, 0);
        let mut chunks = vec![chunk];
        vybe_platform_libc::emitter::string_memory::emit_strchr_linear(&mut chunks, 0, 0);
        chunks[0].emit_op(Op::RETURN, 0);
        let bytes = vybe_platform_wasm::write_wasm_with_options(&chunks,
            vybe_platform_wasm::WasmWriteOptions {
                include_vybe_metadata: false,
                include_name_section: true,
            });
        wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
            .validate_all(&bytes).unwrap();
        let mut vm = vybe_runtime::VM::new();
        assert_eq!(vm.run(chunks).unwrap().as_i32(), expected);
    }
}

#[test]
fn linear_character_search_retains_addresses() {
    let path = Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../tests/c/doom_startup/linear_character_search.c");
    vybe_language_c::register();
    let chunks = vybe_compiler::projects::load(&path).unwrap().compile().unwrap();
    let search = chunks.iter().find(|chunk| chunk.name == "first_match").unwrap();
    let dump = vybe_runtime::debug::disassemble(search);
    assert!(!dump.contains("dyneq"), "dynamic equality in character search: {dump}");
    assert!(!dump.contains("__vybe_slot_"), "dynamic scalar dispatch in character search: {dump}");
    assert!(dump.contains("i32.load8_u"), "linear character search lacks byte loads: {dump}");
    let main = chunks.iter().find(|chunk| chunk.name == "main").unwrap();
    let dump = vybe_runtime::debug::disassemble(main);
    assert!(!dump.contains("dyneq"), "dynamic equality in pointer-returning calls: {dump}");
    let mut vm = vybe_runtime::VM::new();
    vybe_runtime::init_all_registered(&mut vm, &vybe_runtime::capabilities::Capabilities::all());
    vm.run(chunks).unwrap();
    assert_eq!(vm.pending_exit_code, 0);
}

#[test]
fn embedded_command_addresses_keep_record_storage() {
    let path = Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../tests/c/doom_startup/embedded_command_pointer.c");
    vybe_language_c::register();
    let chunks = vybe_compiler::projects::load(&path).unwrap().compile().unwrap();
    let tick = chunks.iter().find(|chunk| chunk.name == "tick").unwrap();
    let dump = vybe_runtime::debug::disassemble(tick);
    assert!(!dump.contains("dyneq"), "dynamic equality in command reads: {dump}");
    assert!(!dump.contains("__vybe_slot_"), "dynamic scalar dispatch in command reads: {dump}");
    assert!(dump.contains("memory.copy"), "command copy lost its bulk-memory path: {dump}");
    assert!(dump.contains("i32.load8_s"), "command byte reads lost sign extension: {dump}");
    assert!(dump.contains("i32.load16_s"), "command word reads lost sign extension: {dump}");
    let mut vm = vybe_runtime::VM::new();
    vybe_runtime::init_all_registered(&mut vm, &vybe_runtime::capabilities::Capabilities::all());
    vm.run(chunks).unwrap();
    assert_eq!(vm.pending_exit_code, 0);
}

#[test]
fn record_union_callback_copies_preserve_callable_fields() {
    let path = Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../tests/c/doom_startup/record_union_callback_copy.c");
    vybe_language_c::register();
    let chunks = vybe_compiler::projects::load(&path).unwrap().compile().unwrap();
    let mut vm = vybe_runtime::VM::new();
    vybe_runtime::init_all_registered(&mut vm, &vybe_runtime::capabilities::Capabilities::all());
    vm.run(chunks).unwrap();
    assert_eq!(vm.pending_exit_code, 0);
}

#[test]
fn scan_counts_and_stream_comparisons_keep_static_types() {
    let path = Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../tests/c/doom_startup/static_scan_stream_results.c");
    vybe_language_c::register();
    let chunks = vybe_compiler::projects::load(&path).unwrap().compile().unwrap();
    for name in ["scan_pair", "is_output", "is_error"] {
        let chunk = chunks.iter().find(|chunk| chunk.name == name).unwrap();
        let dump = vybe_runtime::debug::disassemble(chunk);
        assert!(!dump.contains("dyneq"), "dynamic comparison in {name}: {dump}");
        assert!(!dump.contains("__vybe_slot_"), "numeric comparison-slot dispatch in {name}: {dump}");
    }
    let mut vm = vybe_runtime::VM::new();
    vybe_runtime::init_all_registered(&mut vm, &vybe_runtime::capabilities::Capabilities::all());
    vm.run(chunks).unwrap();
    assert_eq!(vm.pending_exit_code, 0);
}

#[test]
fn callbacks_survive_record_pointer_tables() {
    let path = Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../tests/c/doom_startup/record_callback_table.c");
    vybe_language_c::register();
    let chunks = vybe_compiler::projects::load(&path).unwrap().compile().unwrap();
    let mut vm = vybe_runtime::VM::new();
    vybe_runtime::init_all_registered(&mut vm, &vybe_runtime::capabilities::Capabilities::all());
    vm.run(chunks).unwrap();
    assert_eq!(vm.pending_exit_code, 0);
}

#[test]
fn number_parsing_evaluates_once_without_dynamic_equality() {
    let path = Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../tests/c/doom_startup/number_parse_single_evaluation.c");
    vybe_language_c::register();
    let chunks = vybe_compiler::projects::load(&path).unwrap().compile().unwrap();
    for name in ["decimal", "real_number"] {
        let chunk = chunks.iter().find(|chunk| chunk.name == name).unwrap();
        let dump = vybe_runtime::debug::disassemble(chunk);
        assert!(!dump.contains("dyneq"), "dynamic parser comparison: {dump}");
        assert!(dump.contains("f64.ne"), "missing inline NaN check: {dump}");
    }
    assert!(!chunks.iter().any(|chunk| chunk.name == "__libc_nan_to_default"));
    let mut vm = vybe_runtime::VM::new();
    vybe_runtime::init_all_registered(&mut vm, &vybe_runtime::capabilities::Capabilities::all());
    vm.run(chunks).unwrap();
    assert_eq!(vm.pending_exit_code, 0);
}

#[test]
fn enum_conditionals_preserve_static_comparisons() {
    let path = Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../tests/c/doom_startup/enum_conditional_comparisons.c");
    vybe_language_c::register();
    let chunks = vybe_compiler::projects::load(&path).unwrap().compile().unwrap();
    for name in ["is_original", "is_sequel"] {
        let chunk = chunks.iter().find(|chunk| chunk.name == name).unwrap();
        let dump = vybe_runtime::debug::disassemble(chunk);
        assert!(!dump.contains("dyneq"), "dynamic enum comparison in {name}: {dump}");
    }
    let mut vm = vybe_runtime::VM::new();
    vybe_runtime::init_all_registered(&mut vm, &vybe_runtime::capabilities::Capabilities::all());
    vm.run(chunks).unwrap();
    assert_eq!(vm.pending_exit_code, 0);
}

#[test]
fn menu_character_fields_use_static_comparisons() {
    let path = Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../tests/c/doom_startup/menu_character_fields.c");
    vybe_language_c::register();
    let chunks = vybe_compiler::projects::load(&path).unwrap().compile().unwrap();
    for name in ["matches", "nested_matches"] {
        let chunk = chunks.iter().find(|chunk| chunk.name == name).unwrap();
        let dump = vybe_runtime::debug::disassemble(chunk);
        assert!(!dump.contains("dyneq"), "dynamic character comparison in {name}: {dump}");
    }
    let mut vm = vybe_runtime::VM::new();
    vybe_runtime::init_all_registered(&mut vm, &vybe_runtime::capabilities::Capabilities::all());
    vm.run(chunks).unwrap();
    assert_eq!(vm.pending_exit_code, 0);
}

#[test]
fn gameplay_callback_sentinels_use_static_identity() {
    let path = Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../tests/c/doom_startup/static_gameplay_comparisons.c");
    vybe_language_c::register();
    let chunks = vybe_compiler::projects::load(&path).unwrap().compile().unwrap();
    for name in ["stopped", "nested_stopped", "backwards", "nested_backwards",
        "pressed", "nested_pressed", "removed", "empty_callback"] {
        let chunk = chunks.iter().find(|chunk| chunk.name == name).unwrap();
        let dump = vybe_runtime::debug::disassemble(chunk);
        let lines = dump.lines().collect::<Vec<_>>();
        let offending = lines.iter().position(|line| line.contains("dyneq"));
        if let Some(index) = offending {
            panic!("dynamic equality in {name}: {}",
                lines[index.saturating_sub(8)..(index + 6).min(lines.len())].join("\n"));
        }
    }
    let mut vm = vybe_runtime::VM::new();
    vybe_runtime::init_all_registered(&mut vm, &vybe_runtime::capabilities::Capabilities::all());
    vm.run(chunks).unwrap();
    assert_eq!(vm.pending_exit_code, 0);
}

#[test]
fn byte_cast_keeps_stride_after_conditional_pointer_assignment() {
    let path = Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../tests/c/doom_startup/byte_cast_record_offset.c");
    vybe_language_c::register();
    let chunks = vybe_compiler::projects::load(&path).unwrap().compile().unwrap();
    let mut vm = vybe_runtime::VM::new();
    vybe_runtime::init_all_registered(&mut vm, &vybe_runtime::capabilities::Capabilities::all());
    vm.run(chunks).unwrap_or_else(|error| {
        let frame = error.call_stack.first().unwrap();
        let chunk = vm.chunks.iter().find(|chunk| chunk.name == frame.chunk_name).unwrap();
        let dump = vybe_runtime::debug::disassemble(chunk);
        let window = dump.lines().filter(|line| {
            line.split_whitespace().next().and_then(|text| text.parse::<usize>().ok())
                .is_some_and(|ip| ip >= frame.offset.saturating_sub(200) as usize
                    && ip <= frame.offset as usize + 40)
        }).collect::<Vec<_>>().join("\n");
        panic!("{error}\n{window}");
    });
    assert_eq!(vm.pending_exit_code, 0);
}

#[test]
fn scalar_loop_conditions_do_not_autodereference() {
    let path = Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../tests/c/doom_startup/scalar_loop_truth.c");
    let mut vm = vybe_runtime::VM::new();
    vybe_runtime::init_all_registered(&mut vm, &vybe_runtime::capabilities::Capabilities::all());
    vybe_language_c::register();
    let chunks = vybe_compiler::projects::load(&path).unwrap().compile().unwrap();
    for name in ["countdown", "postcount", "truth"] {
        let helper = chunks.iter().find(|chunk| chunk.name == name).unwrap();
        let dump = vybe_runtime::debug::disassemble(helper);
        assert!(!dump.contains("__vybe_autoderef"), "static condition in {name}: {dump}");
        assert!(!dump.contains("dyneq"), "dynamic equality in {name}: {dump}");
        let condition = dump.split("br_if").next().unwrap();
        assert!(!condition.contains("f64.div"), "loop truth expands into modular conversion in {name}: {dump}");
        if name != "truth" {
            assert!(!dump.contains("f64.div"), "integer decrement is coerced twice: {dump}");
        }
    }
    vm.run(chunks).unwrap();
    assert_eq!(vm.pending_exit_code, 0);
}

#[test]
fn cached_read_workload_vm_only_timing() {
    let path = Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../tests/c/doom_startup/managed_linear_copy_hot.c");
    vybe_language_c::register();
    let absolute = path.canonicalize().unwrap();
    let source = std::fs::read_to_string(&path).unwrap().replace(
        "tests/c/doom_startup/managed_linear_copy_hot.c",
        absolute.to_str().unwrap(),
    );
    let mut bundle = vybe_compiler::projects::load(&path).unwrap();
    bundle.sources[0].code = source;
    let chunks = bundle.compile().unwrap();
    let mut elapsed = Vec::new();
    for _ in 0..5 {
        let mut vm = vybe_runtime::VM::new();
        vybe_runtime::init_all_registered(&mut vm, &vybe_runtime::capabilities::Capabilities::all());
        let program = chunks.clone();
        let start = std::time::Instant::now();
        vm.run(program).unwrap();
        elapsed.push(start.elapsed());
        assert_eq!(vm.pending_exit_code, 0);
    }
    elapsed.sort();
    eprintln!("cached fread workload VM-run median: {:?} (512 x 256 bytes; compilation and VM initialization excluded)", elapsed[2]);
}

#[test]
fn binary_read_transfer_uses_discarded_result_stores() {
    for fixture in ["fread_linear_cached_range.c", "fread_linear_all_byte_values.c"] {
        let path = Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("../../tests/c/doom_startup").join(fixture);
        let mut vm = vybe_runtime::VM::new();
        vybe_runtime::init_all_registered(&mut vm, &vybe_runtime::capabilities::Capabilities::all());
        vybe_language_c::register();
        let chunks = vybe_compiler::projects::load(&path).unwrap().compile().unwrap();
        let helper = chunks.iter().find(|chunk| chunk.name == "__c_fread_into_linear_h").unwrap();
        let copy_index = chunks.iter().position(|chunk| chunk.name == "primitive.byte_copy");
        let mut ip = 0;
        let mut calls_copy = false;
        while ip < helper.code.len() {
            let op = vybe_runtime::Op::from_code_at(&helper.code, ip);
            if op == vybe_runtime::Op::REF_FUNC {
                calls_copy |= copy_index == Some(helper.read_u16(ip + 4) as usize);
            }
            ip += 4 + op.operand_format().size_in(&helper.code, ip + 4);
        }
        let code = if calls_copy {
            &chunks[copy_index.unwrap()].code
        } else {
            &helper.code
        };
        let mut ip = 0;
        let mut stores = 0;
        while ip < code.len() {
            let op = vybe_runtime::Op::from_code_at(code, ip);
            let next = ip + 4 + op.operand_format().size_in(code, ip + 4);
            if op == vybe_runtime::Op::I32_STORE8 {
                stores += 1;
                assert_eq!(vybe_runtime::Op::from_code_at(code, next), vybe_runtime::Op::LOCAL_GET);
                let after_get = next + 4 + vybe_runtime::Op::LOCAL_GET.operand_format().size_in(code, next + 4);
                assert_eq!(vybe_runtime::Op::from_code_at(code, after_get), vybe_runtime::Op::I32_CONST,
                    "byte copy still preserves an unused assignment result");
            }
            ip = next;
        }
        assert!(stores > 0, "missing transfer store in {fixture}");
        vm.run(chunks).unwrap();
        assert_eq!(vm.pending_exit_code, 0, "{fixture}");
    }
}

#[test]
fn record_read_items_use_integer_guards() {
    let path = Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../tests/c/doom_startup/fread_record_items.c");
    vybe_language_c::register();
    let absolute = path.canonicalize().unwrap();
    let mut bundle = vybe_compiler::projects::load(&path).unwrap();
    bundle.sources[0].code = bundle.sources[0].code.replace(
        "tests/c/doom_startup/fread_record_items.c", absolute.to_str().unwrap(),
    );
    let chunks = bundle.compile().unwrap();
    let helper = chunks.iter().find(|chunk| chunk.name == "__c_fread_into_h").unwrap();
    let dump = vybe_runtime::debug::disassemble(helper);
    // The common emitter reverses operands for <= and emits >=.
    assert!(dump.matches("i64.ge_u").count() >= 2, "missing integer size/count guards: {dump}");
    assert!(dump.contains("i64.div_u"), "missing whole-item count: {dump}");
    assert!(!dump.contains("__vybe_slot_"), "dynamic numeric dispatch in fread: {dump}");
    let mut vm = vybe_runtime::VM::new();
    vybe_runtime::init_all_registered(&mut vm, &vybe_runtime::capabilities::Capabilities::all());
    vm.run(chunks).unwrap();
    assert_eq!(vm.pending_exit_code, 0);
}

#[test]
fn pointer_address_lanes_emit_integer_arithmetic() {
    use vybe_ast::{BitLane, Expression, Statement, StmtKind};
    use vybe_compiler::primitives::pointers;
    for (lane, base, index, stride, expected) in [
        (BitLane::W32, 4294967295, 2, 1, "1"),
        (BitLane::W32, 2147483648, 1, 8, "-2147483640"),
        (BitLane::W64, 4294967296, 2, 8, "4294967312"),
    ] {
        let mut module = vybe_language_c::parse("int unused;").unwrap();
        module.body = vec![Statement::new(StmtKind::Return(Some(
            pointers::linear_index_addr_in_lane(Expression::int(base), Expression::int(index), stride, lane),
        )))];
        let profile = vybe_compiler::profile::parse_profile(vybe_language_c::profile_source()).unwrap();
        let chunks = vybe_compiler::Compiler::with_profile(profile).compile(&module).unwrap();
        let dump = vybe_runtime::debug::disassemble(&chunks[0]);
        let code = &chunks[0].code;
        let mut ip = 0;
        let mut narrowings = 0;
        while ip < code.len() {
            let op = vybe_runtime::Op::from_code_at(code, ip);
            assert!(op != vybe_runtime::Op::F64_ADD && op != vybe_runtime::Op::F64_MUL,
                "floating address arithmetic: {dump}");
            if op == vybe_runtime::Op::RETURN { break; }
            if op == vybe_runtime::Op::I64_TRUNC_F64_S { narrowings += 1; }
            ip += 4 + op.operand_format().size_in(code, ip + 4);
        }
        if lane == BitLane::W32 && stride == 8 {
            assert!(narrowings <= 1, "intermediate i32 address was narrowed again: {dump}");
        }
        let mut vm = vybe_runtime::VM::new();
        vybe_runtime::init_all_registered(&mut vm, &vybe_runtime::capabilities::Capabilities::all());
        assert_eq!(vm.run(chunks).unwrap().to_string(), expected);
    }
}

#[test]
fn shared_binding_directives_control_initializer_and_assignment_conversion() {
    use vybe_ast::{BindingPattern, ExprKind, Expression, Literal, Statement, StmtKind,
        TypeHint, VarDeclarator, VarDeclKind};
    for (hint, first, next, expected_first, expected_next, converts) in [
        (TypeHint::checked("uint32"), 4294967295.0, 7.0, "4294967295", "7", false),
        (TypeHint::converting("uint32"), 23.75, 7.75, "23", "7", true),
        (TypeHint::descriptive("uint32"), 23.75, 7.75, "23.75", "7.75", false),
    ] {
        for assign in [false, true] {
            let mut module = vybe_language_c::parse("int unused;").unwrap();
            let id = || Expression::new(ExprKind::Ident("value".into()));
            let number = |value| Expression::new(ExprKind::Lit(Literal::Float(value)));
            module.body = vec![Statement::new(StmtKind::VarDecl {
                declarations: vec![VarDeclarator {
                    pattern: BindingPattern::Ident("value".into()),
                    type_hint: Some(hint.clone()),
                    init: Some(number(first)), array_bounds: None, with_events: false,
                }], kind: VarDeclKind::Var,
            })];
            if assign {
                module.body.push(Statement::new(StmtKind::Expr(Expression::new(ExprKind::Assign {
                    target: Box::new(id()), value: Box::new(number(next)),
                }))));
            }
            module.body.push(Statement::new(StmtKind::Return(Some(id()))));
            let profile = vybe_compiler::profile::parse_profile(vybe_language_c::profile_source()).unwrap();
            let chunks = vybe_compiler::Compiler::with_profile(profile).compile(&module).unwrap();
            if !converts {
                let dump = vybe_runtime::debug::disassemble(&chunks[0]);
                assert!(!dump.contains("f64.div"), "binding {:?}: {dump}", hint.binding);
            }
            let mut vm = vybe_runtime::VM::new();
            vybe_runtime::init_all_registered(&mut vm, &vybe_runtime::capabilities::Capabilities::all());
            let value = vm.run(chunks).unwrap();
            assert_eq!(value.to_string(), if assign { expected_next } else { expected_first },
                "binding {:?}, assignment={assign}", hint.binding);
        }
    }
}

#[test]
fn unsigned_span_arithmetic_does_not_repeat_binding_conversions() {
    let path = Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../tests/c/doom_startup/span_unsigned_math.c");
    let mut vm = vybe_runtime::VM::new();
    vybe_runtime::init_all_registered(
        &mut vm,
        &vybe_runtime::capabilities::Capabilities::all(),
    );
    vybe_language_c::register();
    let chunks = vybe_compiler::projects::load(&path).unwrap().compile().unwrap();
    let span = chunks.iter().find(|chunk| chunk.name == "span_math").unwrap();
    let dump = vybe_runtime::debug::disassemble(span);
    let word = chunks.iter().find(|chunk| chunk.name == "reinterpret_word").unwrap();
    let word_dump = vybe_runtime::debug::disassemble(word);
    vm.run(chunks.clone()).unwrap();
    assert_eq!(vm.pending_exit_code, 0);
    assert!(!dump.contains("f64.div"), "redundant unsigned binding conversion: {dump}");
    assert!(!dump.contains("dyneq"), "dynamic span equality: {dump}");
    assert!(!word_dump.contains("f64.div"), "redundant integer word conversion: {word_dump}");
    println!("span_math bytecode bytes: {}", span.code.len());
}

#[test]
fn record_array_updates_preserve_the_field_place() {
    let path = Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../tests/c/doom_startup/button_state.c");
    let mut vm = vybe_runtime::VM::new();
    vybe_runtime::init_all_registered(
        &mut vm,
        &vybe_runtime::capabilities::Capabilities::all(),
    );
    vybe_language_c::register();
    let chunks = vybe_compiler::projects::load(&path).unwrap().compile().unwrap();
    let main = chunks.iter().find(|chunk| chunk.name == "main").unwrap();
    let dump = vybe_runtime::debug::disassemble(main);
    assert!(!dump.contains("dyneq"), "dynamic record-array comparison: {dump}");
    vm.run(chunks).unwrap();
    assert_eq!(vm.pending_exit_code, 0);
}

#[test]
fn unsigned_angles_convert_when_assigned_to_signed_scalars() {
    let path = Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../tests/c/doom_startup/wall_angle_distance.c");
    let mut vm = vybe_runtime::VM::new();
    vybe_runtime::init_all_registered(
        &mut vm,
        &vybe_runtime::capabilities::Capabilities::all(),
    );
    vybe_language_c::register();
    let chunks = vybe_compiler::projects::load(&path).unwrap().compile().unwrap();
    vm.run(chunks.clone()).unwrap();
    assert_eq!(vm.pending_exit_code, 0);
    for name in ["offset", "distance", "main", "local_assignments",
        "compound_assignment", "array_assignment", "record_assignment"] {
        let chunk = chunks.iter().find(|chunk| chunk.name == name).unwrap();
        let dump = vybe_runtime::debug::disassemble(chunk);
        let dynamic = dump.lines().filter(|line| line.contains("dyneq"))
            .collect::<Vec<_>>().join("\n");
        assert!(dynamic.is_empty(), "dynamic angle comparison in {name}: {dynamic}");
    }
}

#[test]
fn texture_pointer_tables_preserve_record_and_scalar_backing() {
    let path = Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../tests/c/doom_startup/texture_pointer_tables.c");
    let mut vm = vybe_runtime::VM::new();
    vybe_runtime::init_all_registered(
        &mut vm,
        &vybe_runtime::capabilities::Capabilities::all(),
    );
    vybe_language_c::register();
    let chunks = vybe_compiler::projects::load(&path).unwrap().compile().unwrap();
    let main = chunks.iter().find(|chunk| chunk.name == "main").unwrap();
    let dump = vybe_runtime::debug::disassemble(main);
    vm.run(chunks).unwrap_or_else(|error| {
        let offset = error.call_stack.first().map(|frame| frame.offset as usize).unwrap_or(0);
        let window = dump.lines().filter(|line| {
            line.split_whitespace().next().and_then(|text| text.parse::<usize>().ok())
                .is_some_and(|ip| ip >= offset.saturating_sub(1600) && ip <= offset)
        }).collect::<Vec<_>>().join("\n");
        panic!("{error}\n{window}");
    });
    assert_eq!(vm.pending_exit_code, 0);
}

#[test]
fn doom_projection_matches_native_tables() {
    let path = Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../tests/c/doom_startup/doom_projection_tables.c");
    let mut vm = vybe_runtime::VM::new();
    vybe_runtime::init_all_registered(
        &mut vm,
        &vybe_runtime::capabilities::Capabilities::all(),
    );
    vybe_language_c::register();
    let chunks = vybe_compiler::projects::load(&path).unwrap().compile().unwrap();
    for name in ["mapping", "scale", "multiply", "divide", "main"] {
        let chunk = chunks.iter().find(|chunk| chunk.name == name).unwrap();
        let dump = vybe_runtime::debug::disassemble(chunk);
        assert!(!dump.contains("dyneq"), "dynamic projection comparison in {name}: {dump}");
        if name == "divide" {
            assert!(!dump.contains("__vybe_slot_"), "dynamic absolute-value comparison: {dump}");
        }
    }
    vm.run(chunks).unwrap();
    assert_eq!(vm.pending_exit_code, 0);
}

#[test]
fn scoped_embedded_patch_pointers_keep_backing() {
    let path = Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../tests/c/doom_startup/scoped_embedded_patch_pointer.c");
    let mut vm = vybe_runtime::VM::new();
    vybe_runtime::init_all_registered(
        &mut vm,
        &vybe_runtime::capabilities::Capabilities::all(),
    );
    vybe_language_c::register();
    let chunks = vybe_compiler::projects::load(&path).unwrap().compile().unwrap();
    let dump = vybe_runtime::debug::disassemble(chunks.iter().find(|chunk| chunk.name == "main").unwrap());
    vm.run(chunks).unwrap_or_else(|error| {
        let offset = error.call_stack.first().map(|frame| frame.offset as usize).unwrap_or(0);
        let window = dump.lines().filter(|line| {
            line.split_whitespace().next().and_then(|text| text.parse::<usize>().ok())
                .is_some_and(|ip| ip >= offset.saturating_sub(600) && ip <= offset + 40)
        }).collect::<Vec<_>>().join("\n");
        panic!("{error}\n{window}");
    });
    assert_eq!(vm.pending_exit_code, 0);
}

#[test]
fn mixed_word_pointer_offsets_preserve_both_backings() {
    let path = Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../tests/c/doom_startup/mixed_word_pointer_offsets.c");
    vybe_language_c::register();
    let chunks = vybe_compiler::projects::load(&path).unwrap().compile().unwrap();
    let bytes = vybe_platform_wasm::write_wasm_with_options(&chunks,
        vybe_platform_wasm::WasmWriteOptions {
            include_vybe_metadata: false,
            include_name_section: true,
        });
    validate_wasm_with_context(&bytes, &chunks);
    assert_no_dynamic_equality(&chunks);
    let mut vm = vybe_runtime::VM::new();
    vybe_runtime::init_all_registered(&mut vm, &vybe_runtime::capabilities::Capabilities::all());
    vm.run(chunks).unwrap();
    assert_eq!(vm.pending_exit_code, 0);
}

#[test]
fn span_frame_preserves_full_width_wrapped_texture_pixels() {
    let path = Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../tests/c/doom_startup/doom_span_frame.c");
    vybe_language_c::register();
    let chunks = vybe_compiler::projects::load(&path).unwrap().compile().unwrap();
    let bytes = vybe_platform_wasm::write_wasm_with_options(&chunks,
        vybe_platform_wasm::WasmWriteOptions {
            include_vybe_metadata: false,
            include_name_section: true,
        });
    validate_wasm_with_context(&bytes, &chunks);
    assert_no_dynamic_equality(&chunks);
    let mut vm = vybe_runtime::VM::new();
    vybe_runtime::init_all_registered(&mut vm, &vybe_runtime::capabilities::Capabilities::all());
    vm.run(chunks).unwrap();
    assert_eq!(vm.pending_exit_code, 0);
}

#[test]
fn record_word64_fields_cross_value_branches_with_declared_signedness() {
    let path = Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../tests/c/doom_startup/record_word64_value.c");
    vybe_language_c::register();
    let chunks = vybe_compiler::projects::load(&path).unwrap().compile().unwrap();
    let bytes = vybe_platform_wasm::write_wasm_with_options(&chunks,
        vybe_platform_wasm::WasmWriteOptions { include_vybe_metadata: false, include_name_section: true });
    validate_wasm_with_context(&bytes, &chunks);
    assert_no_dynamic_equality(&chunks);
    let mut vm = vybe_runtime::VM::new();
    vybe_runtime::init_all_registered(&mut vm, &vybe_runtime::capabilities::Capabilities::all());
    vm.run(chunks).unwrap();
    assert_eq!(vm.pending_exit_code, 0);
}

#[test]
fn static_counter_postincrement_preserves_value_across_reference_store() {
    let path = Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../tests/c/doom_startup/static_counter_value.c");
    vybe_language_c::register();
    let chunks = vybe_compiler::projects::load(&path).unwrap().compile().unwrap();
    let bytes = vybe_platform_wasm::write_wasm_with_options(&chunks,
        vybe_platform_wasm::WasmWriteOptions { include_vybe_metadata: false, include_name_section: true });
    validate_wasm_with_context(&bytes, &chunks);
    assert_no_dynamic_equality(&chunks);
    let mut vm = vybe_runtime::VM::new();
    vybe_runtime::init_all_registered(&mut vm, &vybe_runtime::capabilities::Capabilities::all());
    vm.run(chunks).unwrap();
    assert_eq!(vm.pending_exit_code, 0);
}

#[test]
fn column_buffers_preserve_mixed_backings() {
    let path = Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../tests/c/doom_startup/doom_column_fractional_draw.c");
    let mut vm = vybe_runtime::VM::new();
    vybe_runtime::init_all_registered(
        &mut vm,
        &vybe_runtime::capabilities::Capabilities::all(),
    );
    vybe_language_c::register();
    let chunks = vybe_compiler::projects::load(&path).unwrap().compile().unwrap();
    vm.run(chunks).unwrap();
    assert_eq!(vm.pending_exit_code, 0);
}

#[test]
fn integral_switches_have_static_case_comparisons() {
    let path = Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../tests/c/doom_startup/static_integral_switch.c");
    let mut vm = vybe_runtime::VM::new();
    vybe_runtime::init_all_registered(
        &mut vm,
        &vybe_runtime::capabilities::Capabilities::all(),
    );
    vybe_language_c::register();
    let chunks = vybe_compiler::projects::load(&path).unwrap().compile().unwrap();
    for name in ["direction", "token", "promoted", "selected", "unsigned_case", "ranged"] {
        let chunk = chunks.iter().find(|chunk| chunk.name == name).unwrap();
        let dump = vybe_runtime::debug::disassemble(chunk);
        assert!(!dump.contains("dyneq"), "dynamic switch comparison in {name}: {dump}");
        println!("{name} switch bytecode bytes: {}", chunk.code.len());
    }
    vm.run(chunks).unwrap();
    assert_eq!(vm.pending_exit_code, 0);
}

#[test]
fn sdl_events_roundtrip_between_record_backings() {
    let path = Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../tests/c/doom_startup/sdl_keyboard_translation_table.c");
    let mut vm = vybe_runtime::VM::new();
    vybe_runtime::init_all_registered(
        &mut vm,
        &vybe_runtime::capabilities::Capabilities::all(),
    );
    vybe_language_c::register();
    let chunks = vybe_compiler::projects::load(&path).unwrap().compile().unwrap();
    for name in ["window_matches", "renderer_failed", "output_failed", "peeked",
        "window_fullscreen", "display_failed", "bounds_failed", "fullscreen_shortcut",
        "ascii_text", "__libc_sdl_input"] {
        let helper = chunks.iter().find(|chunk| chunk.name == name).unwrap();
        let dump = vybe_runtime::debug::disassemble(helper);
        assert!(!dump.contains("dyneq"), "dynamic SDL comparison in {name}: {dump}");
    }
    vm.run(chunks).unwrap();
    assert_eq!(vm.pending_exit_code, 0);
}

#[test]
fn sdl_peeping_preserves_filters_and_queue_order() {
    let path = Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../tests/c/doom_startup/sdl_peep_event_queue.c");
    let mut vm = vybe_runtime::VM::new();
    vybe_runtime::init_all_registered(
        &mut vm, &vybe_runtime::capabilities::Capabilities::all(),
    );
    vybe_language_c::register();
    let chunks = vybe_compiler::projects::load(&path).unwrap().compile().unwrap();
    let main = chunks.iter().find(|chunk| chunk.name == "main").unwrap();
    let dump = vybe_runtime::debug::disassemble(main);
    assert!(!dump.contains("dyneq"), "dynamic queue equality: {dump}");
    vm.run(chunks).unwrap();
    assert_eq!(vm.pending_exit_code, 0);
}

#[test]
fn sdl_queue_wasm_reload_runs_without_metadata() {
    let path = Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../tests/c/doom_startup/sdl_peep_event_queue.c");
    let mut vm = vybe_runtime::VM::new();
    vybe_runtime::init_all_registered(
        &mut vm, &vybe_runtime::capabilities::Capabilities::all(),
    );
    vybe_language_c::register();
    let chunks = vybe_compiler::projects::load(&path).unwrap().compile().unwrap();
    let wasm = vybe_platform_wasm::write_wasm_with_options(
        &chunks,
        vybe_platform_wasm::WasmWriteOptions {
            include_vybe_metadata: false,
            include_name_section: true,
        },
    );
    if let Some(path) = std::env::var_os("VYBE_SDL_RELOAD_ARTIFACT") {
        std::fs::write(path, &wasm).expect("save requested SDL reload artifact");
    }
    wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
        .validate_all(&wasm).unwrap();
    let reloaded = vybe_platform_wasm::read_wasm(&wasm).unwrap();
    let diagnostic = reloaded.clone();
    let strings = vybe_platform_wasm::writer::sections::collect_string_constants(&chunks);
    let hosts = vybe_platform_wasm::writer::sections::collect_host_globals(&chunks);
    let defined = vybe_platform_wasm::writer::sections::collect_globals(&chunks, &strings, &hosts);
    let global_names = vybe_runtime::chunk::global_index_space(&strings, &hosts, &defined);
    vm.run(reloaded).unwrap_or_else(|error| {
        if let Some(frame) = error.call_stack.first() {
            let offset = frame.offset as usize;
            for chunk in diagnostic.iter().filter(|chunk|
                chunk.name == frame.chunk_name && offset < chunk.code.len()) {
                let mut ip = 0;
                while ip < chunk.code.len() {
                    let (instruction, next) = vybe_runtime::debug::disassemble_instruction(chunk, ip);
                    if ip >= offset.saturating_sub(160) && ip <= offset + 40 {
                        eprintln!("{instruction}");
                        if instruction.starts_with("global.get ") {
                            let index = chunk.read_u32(ip + 4);
                            if let Some(name) = chunk.globals.get(index as usize) {
                                eprintln!("global value {name}: {:?}", vm.global(name));
                            }
                            if let Some(name) = global_names.get(index as usize) {
                                eprintln!("compiler binding: {name}");
                                if let Some(root) = chunks.first() {
                                    let dump = vybe_runtime::debug::disassemble(root);
                                    let lines: Vec<_> = dump.lines().collect();
                                    for (line, text) in lines.iter().enumerate()
                                        .filter(|(_, text)| text.contains(&format!("({name})"))) {
                                        eprintln!("binding source: {text}");
                                        for nearby in &lines[line.saturating_sub(4)..(line + 5).min(lines.len())] {
                                            eprintln!("{nearby}");
                                        }
                                    }
                                }
                                for init in chunks.iter().flat_map(|c| &c.global_inits)
                                    .filter(|init| &init.name == name) {
                                    eprintln!("binding initializer: {init:?}");
                                }
                            }
                        }
                    }
                    ip = next;
                }
            }
        }
        panic!("{error}");
    });
    assert_eq!(vm.pending_exit_code, 0);
}

#[test]
fn pointer_distances_have_static_comparisons() {
    let path = Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../tests/c/doom_startup/static_pointer_distance.c");
    let mut vm = vybe_runtime::VM::new();
    vybe_runtime::init_all_registered(
        &mut vm,
        &vybe_runtime::capabilities::Capabilities::all(),
    );
    vybe_language_c::register();
    let chunks = vybe_compiler::projects::load(&path).unwrap().compile().unwrap();
    let bytes = vybe_platform_wasm::write_wasm_with_options(&chunks,
        vybe_platform_wasm::WasmWriteOptions {
            include_vybe_metadata: false,
            include_name_section: true,
        });
    wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
        .validate_all(&bytes).unwrap();
    for name in ["full", "before", "check_openings"] {
        let chunk = chunks.iter().find(|chunk| chunk.name == name).unwrap();
        let dump = vybe_runtime::debug::disassemble(chunk);
        assert!(!dump.contains("dyneq"), "dynamic pointer distance: {dump}");
    }
    vm.run(chunks).unwrap();
    assert_eq!(vm.pending_exit_code, 0);
}

#[test]
fn save_menu_strings_copy_across_pointer_backings() {
    let path = Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../tests/c/doom_startup/save_menu_string_copy.c");
    vybe_language_c::register();
    let mut vm = vybe_runtime::VM::new();
    vybe_runtime::init_all_registered(&mut vm, &vybe_runtime::capabilities::Capabilities::all());
    let compiled = vybe_compiler::dynamic::RuntimeCompilerService::new(&mut vm)
        .compile_path(&path).unwrap();
    for chunk in &compiled.chunks {
        let dump = vybe_runtime::debug::disassemble(chunk);
        assert!(!dump.contains("dyneq"),
            "dynamic save-menu equality in {}: {dump}", chunk.name);
    }
    let bytes = vybe_platform_wasm::write_wasm_with_options(&compiled.chunks,
        vybe_platform_wasm::WasmWriteOptions {
            include_vybe_metadata: false,
            include_name_section: true,
        });
    validate_wasm_with_context(&bytes, &compiled.chunks);
    if std::env::var_os("VYBE_C_SAVE_MENU_DUMP").is_some() {
        for chunk in &compiled.chunks {
            let dump = vybe_runtime::debug::disassemble(chunk);
            if matches!(chunk.name.as_str(), "save_file" | "__c_write_carray_string"
                | "main" | "worker" | "__c_thread_entry")
                || (chunk.name == "<lambda>" && dump.contains("__c_write_carray_string"))
            {
                println!("{dump}");
            }
        }
    }
    let result = vybe_compiler::dynamic::RuntimeCompilerService::new(&mut vm)
        .run_compiled(compiled);
    let stage = vm.globals_by_name().into_iter()
        .find(|(name, _)| name == "copy_stage");
    assert!(result.is_ok(), "save-menu stage {stage:?}: {result:?}");
    assert_eq!(vm.pending_exit_code, 0);
}

#[test]
fn deh_literal_keys_keep_bytes_in_zone_storage() {
    let path = Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../tests/c/doom_startup/deh_zone_replacements.c");
    vybe_language_c::register();
    let mut vm = vybe_runtime::VM::new();
    vybe_runtime::init_all_registered(&mut vm, &vybe_runtime::capabilities::Capabilities::all());
    let compiled = vybe_compiler::dynamic::RuntimeCompilerService::new(&mut vm)
        .compile_path(&path).unwrap();
    assert_no_dynamic_equality(&compiled.chunks);
    let bytes = vybe_platform_wasm::write_wasm_with_options(&compiled.chunks,
        vybe_platform_wasm::WasmWriteOptions {
            include_vybe_metadata: false,
            include_name_section: true,
        });
    validate_wasm_with_context(&bytes, &compiled.chunks);
    vybe_compiler::dynamic::RuntimeCompilerService::new(&mut vm)
        .run_compiled(compiled).unwrap();
    assert_eq!(vm.pending_exit_code, 0);
}

#[test]
fn automap_nested_endpoints_do_not_expand_record_owners() {
    let path = Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../tests/c/doom_startup/automap_nested_endpoints.c");
    vybe_language_c::register();
    let mut vm = vybe_runtime::VM::new();
    vybe_runtime::init_all_registered(&mut vm, &vybe_runtime::capabilities::Capabilities::all());
    let compiled = vybe_compiler::dynamic::RuntimeCompilerService::new(&mut vm)
        .compile_path(&path).unwrap();
    let draw = compiled.chunks.iter().find(|chunk| chunk.name == "draw_character").unwrap();
    eprintln!("draw_character: {} bytecode bytes, {} locals", draw.code.len(), draw.local_count);
    assert!(draw.code.len() < 1_000_000, "nested endpoint lowering expanded");
    assert_no_dynamic_equality(&compiled.chunks);
    let bytes = vybe_platform_wasm::write_wasm_with_options(&compiled.chunks,
        vybe_platform_wasm::WasmWriteOptions {
            include_vybe_metadata: false,
            include_name_section: true,
        });
    validate_wasm_with_context(&bytes, &compiled.chunks);
    vybe_compiler::dynamic::RuntimeCompilerService::new(&mut vm)
        .run_compiled(compiled).unwrap();
    assert_eq!(vm.pending_exit_code, 0);
}

#[test]
fn zero_record_arrays_use_distinct_bounded_initializers() {
    let path = Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../tests/c/doom_startup/zero_nested_record_arrays.c");
    vybe_language_c::register();
    let mut vm = vybe_runtime::VM::new();
    vybe_runtime::init_all_registered(&mut vm, &vybe_runtime::capabilities::Capabilities::all());
    let compiled = vybe_compiler::dynamic::RuntimeCompilerService::new(&mut vm)
        .compile_path(&path).unwrap();
    for chunk in &compiled.chunks {
        assert!(chunk.code.len() < 200_000,
            "record array expanded in {}: {} bytes", chunk.name, chunk.code.len());
    }
    assert_no_dynamic_equality(&compiled.chunks);
    let bytes = vybe_platform_wasm::write_wasm_with_options(&compiled.chunks,
        vybe_platform_wasm::WasmWriteOptions {
            include_vybe_metadata: false,
            include_name_section: true,
        });
    validate_wasm_with_context(&bytes, &compiled.chunks);
    vybe_compiler::dynamic::RuntimeCompilerService::new(&mut vm)
        .run_compiled(compiled).unwrap();
    assert_eq!(vm.pending_exit_code, 0);
}

#[test]
fn nested_record_values_materialize_managed_and_linear_fields() {
    let path = Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../tests/c/doom_startup/nested_record_value_copy.c");
    vybe_language_c::register();
    let mut vm = vybe_runtime::VM::new();
    vybe_runtime::init_all_registered(&mut vm, &vybe_runtime::capabilities::Capabilities::all());
    let compiled = vybe_compiler::dynamic::RuntimeCompilerService::new(&mut vm)
        .compile_path(&path).unwrap();
    assert_no_dynamic_equality(&compiled.chunks);
    let bytes = vybe_platform_wasm::write_wasm_with_options(&compiled.chunks,
        vybe_platform_wasm::WasmWriteOptions {
            include_vybe_metadata: false,
            include_name_section: true,
        });
    validate_wasm_with_context(&bytes, &compiled.chunks);
    vybe_compiler::dynamic::RuntimeCompilerService::new(&mut vm)
        .run_compiled(compiled).unwrap();
    assert_eq!(vm.pending_exit_code, 0);
}

#[test]
fn threaded_player_movement_keeps_record_type_mapping() {
    let path = Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../tests/c/doom_startup/threaded_player_movement.c");
    vybe_language_c::register();
    let mut vm = vybe_runtime::VM::new();
    vybe_runtime::init_all_registered(&mut vm, &vybe_runtime::capabilities::Capabilities::all());
    let compiled = vybe_compiler::dynamic::RuntimeCompilerService::new(&mut vm)
        .compile_path(&path).unwrap();
    for chunk in &compiled.chunks {
        assert!(!vybe_runtime::debug::disassemble(chunk).contains("dyneq"),
            "dynamic threaded movement equality in {}", chunk.name);
    }
    let bytes = vybe_platform_wasm::write_wasm_with_options(&compiled.chunks,
        vybe_platform_wasm::WasmWriteOptions {
            include_vybe_metadata: false,
            include_name_section: true,
        });
    validate_wasm_with_context(&bytes, &compiled.chunks);
    vybe_compiler::dynamic::RuntimeCompilerService::new(&mut vm)
        .run_compiled(compiled).unwrap();
    assert_eq!(vm.pending_exit_code, 0);
}

#[test]
fn intercept_traversal_preserves_flags_union_and_callback_argument() {
    let path = Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../tests/c/doom_startup/intercept_traversal.c");
    vybe_language_c::register();
    let mut vm = vybe_runtime::VM::new();
    vybe_runtime::init_all_registered(&mut vm, &vybe_runtime::capabilities::Capabilities::all());
    let compiled = vybe_compiler::dynamic::RuntimeCompilerService::new(&mut vm)
        .compile_path(&path).unwrap();
    for chunk in &compiled.chunks {
        assert!(!vybe_runtime::debug::disassemble(chunk).contains("dyneq"),
            "dynamic intercept comparison in {}", chunk.name);
    }
    let bytes = vybe_platform_wasm::write_wasm_with_options(&compiled.chunks,
        vybe_platform_wasm::WasmWriteOptions {
            include_vybe_metadata: false,
            include_name_section: true,
        });
    validate_wasm_with_context(&bytes, &compiled.chunks);
    vybe_compiler::dynamic::RuntimeCompilerService::new(&mut vm)
        .run_compiled(compiled).unwrap();
    assert_eq!(vm.pending_exit_code, 0);
}

#[test]
fn player_movement_survives_runtime_compiler_service() {
    let path = Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../tests/c/doom_startup/player_thinker_movement.c");
    vybe_language_c::register();
    let mut vm = vybe_runtime::VM::new();
    vybe_runtime::init_all_registered(&mut vm, &vybe_runtime::capabilities::Capabilities::all());
    let compiled = vybe_compiler::dynamic::RuntimeCompilerService::new(&mut vm)
        .compile_path(&path).unwrap();
    for chunk in &compiled.chunks {
        assert!(!vybe_runtime::debug::disassemble(chunk).contains("dyneq"),
            "dynamic movement equality in {}", chunk.name);
    }
    vybe_compiler::dynamic::RuntimeCompilerService::new(&mut vm)
        .run_compiled(compiled).unwrap();
    assert_eq!(vm.pending_exit_code, 0);
    for (name, expected) in [("calls", 4), ("destination_evaluations", 4),
        ("source_evaluations", 4), ("observed_forward", 10),
        ("observed_momentum", 20480), ("observed_x", -12511232)] {
        assert_eq!(vm.global(name).expect("movement result must exist").as_i32(),
            expected, "incorrect final movement result: {name}");
    }
}

#[test]
fn player_commands_reach_heap_thinker_movement() {
    let path = Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../tests/c/doom_startup/player_thinker_movement.c");
    vybe_language_c::register();
    let chunks = vybe_compiler::projects::load(&path).unwrap().compile().unwrap();
    let bytes = vybe_platform_wasm::write_wasm_with_options(&chunks,
        vybe_platform_wasm::WasmWriteOptions {
            include_vybe_metadata: false,
            include_name_section: true,
        });
    wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
        .validate_all(&bytes).unwrap();
    for name in ["main", "tick", "movement", "add", "initialize_direct_list",
        "add_direct_node", "run_direct_thinkers",
        "check_direct_movement"] {
        let chunk = chunks.iter().find(|chunk| chunk.name == name).unwrap();
        let dump = vybe_runtime::debug::disassemble(chunk);
        assert!(!dump.contains("dyneq"), "dynamic movement comparison: {dump}");
    }
    let mut vm = vybe_runtime::VM::new();
    vybe_runtime::init_all_registered(&mut vm, &vybe_runtime::capabilities::Capabilities::all());
    vm.run(chunks).unwrap();
    assert_eq!(vm.pending_exit_code, 0);
}

#[test]
fn global_address_promotion_survives_function_compilation_order() {
    use vybe_ast::{BindingPattern, ExprKind, Expression, Modifiers, PlaceExpr,
        Statement, StmtKind, VarDeclarator, VarDeclKind};
    vybe_language_c::register();
    let address = || Expression::new(ExprKind::RefOf(Box::new(
        PlaceExpr::Ident("value".into()),
    )));
    let declaration = |name: &str, init: Expression| Statement::new(StmtKind::VarDecl {
        declarations: vec![VarDeclarator {
            pattern: BindingPattern::Ident(name.into()), type_hint: None,
            init: Some(init), array_bounds: None, with_events: false,
        }], kind: VarDeclKind::Var,
    });
    for function_first in [false, true] {
        let mut module = vybe_language_c::parse("int unused;").unwrap();
        let function = Statement::new(StmtKind::FunctionDecl {
            name: "address".into(), params: vec![], return_type: None,
            body: vec![Statement::new(StmtKind::Return(Some(address())))],
            modifiers: Modifiers::default(), handles: vec![], is_async: false,
            is_generator: false, is_sub: false,
        });
        module.body = vec![declaration("root", address()),
            declaration("value", Expression::int(17))];
        module.body.insert(if function_first { 0 } else { 2 }, function);
        module.body.push(Statement::new(StmtKind::Return(Some(
            Expression::new(ExprKind::RefLoad(Box::new(Expression::ident("root")))),
        ))));
        let profile = vybe_compiler::profile::parse_profile(
            vybe_language_c::profile_source(),
        ).unwrap();
        let chunks = vybe_compiler::Compiler::with_profile(profile).compile(&module).unwrap();
        let bytes = vybe_platform_wasm::write_wasm_with_options(&chunks,
            vybe_platform_wasm::WasmWriteOptions {
                include_vybe_metadata: false, include_name_section: true,
            });
        wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
            .validate_all(&bytes).unwrap();
        let mut vm = vybe_runtime::VM::new();
        vybe_runtime::init_all_registered(&mut vm, &vybe_runtime::capabilities::Capabilities::all());
        assert_eq!(vm.run(chunks).unwrap().as_f64(), 17.0);
    }
}

#[test]
fn keyboard_bindings_build_static_signed_commands() {
    let path = Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../tests/c/doom_startup/keyboard_command_input.c");
    vybe_language_c::register();
    let chunks = vybe_compiler::projects::load(&path).unwrap().compile().unwrap();
    let bytes = vybe_platform_wasm::write_wasm_with_options(&chunks,
        vybe_platform_wasm::WasmWriteOptions {
            include_vybe_metadata: false,
            include_name_section: true,
        });
    wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
        .validate_all(&bytes).unwrap();
    for name in ["main", "build", "bind_key"] {
        let chunk = chunks.iter().find(|chunk| chunk.name == name).unwrap();
        let dump = vybe_runtime::debug::disassemble(chunk);
        assert!(!dump.contains("dyneq"), "dynamic keyboard command comparison: {dump}");
    }
    let mut vm = vybe_runtime::VM::new();
    vybe_runtime::init_all_registered(&mut vm, &vybe_runtime::capabilities::Capabilities::all());
    vm.run(chunks).unwrap();
    assert_eq!(vm.pending_exit_code, 0);
}

#[test]
fn pointer_ordering_uses_backing_coordinates_once() {
    let path = Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../tests/c/doom_startup/static_pointer_ordering.c");
    vybe_language_c::register();
    let chunks = vybe_compiler::projects::load(&path).unwrap().compile().unwrap();
    assert_no_dynamic_equality(&chunks);
    for name in ["after", "ordered", "record_after", "evaluated_once"] {
        let dump = function_dependency_dump(&chunks, name);
        assert!(!dump.contains("dyneq"), "dynamic pointer equality in {name}: {dump}");
        for helper in ["dynlt", "dynle", "dyngt", "dynge"] {
            assert!(!dump.contains(helper), "dynamic pointer ordering in {name}: {dump}");
        }
        assert!(!dump.contains("__vybe_slot_"), "dynamic pointer ordering in {name}: {dump}");
        assert!(dump.contains("i64.lt_s") || dump.contains("i64.gt_s"),
            "missing signed integer pointer ordering in {name}: {dump}");
    }
    let dump = function_dependency_dump(&chunks, "linear_after");
    assert!(dump.contains("i32.gt_u"), "missing direct linear comparison: {dump}");
    assert!(!dump.contains("__byte_offset"), "linear comparison uses managed coordinates: {dump}");
    let wasm = vybe_platform_wasm::write_wasm_with_options(
        &chunks,
        vybe_platform_wasm::WasmWriteOptions {
            include_vybe_metadata: false,
            include_name_section: true,
        },
    );
    wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
        .validate_all(&wasm).unwrap();
    let mut vm = vybe_runtime::VM::new();
    vybe_runtime::init_all_registered(&mut vm, &vybe_runtime::capabilities::Capabilities::all());
    vm.run(chunks).unwrap();
    assert_eq!(vm.pending_exit_code, 0);
}

fn function_dependency_dump(chunks: &[vybe_runtime::Chunk], root: &str) -> String {
    use std::collections::{HashMap, HashSet};
    use vybe_runtime::Op;
    let functions: HashMap<_, _> = chunks.iter().enumerate()
        .map(|(index, chunk)| (chunk.name.as_str(), index)).collect();
    let mut pending = vec![*functions.get(root).expect("compiled function")];
    let mut visited = HashSet::new();
    let mut dumps = Vec::new();
    while let Some(index) = pending.pop() {
        if !visited.insert(index) { continue; }
        let chunk = &chunks[index];
        dumps.push(vybe_runtime::debug::disassemble(chunk));
        let mut offset = 0;
        while offset < chunk.code.len() {
            match Op::decode(chunk.read_u16(offset), chunk.read_u16(offset + 2)).unwrap() {
                Op::GLOBAL_GET => {
                    let globals = if chunk.globals.is_empty() {
                        &chunks[0].globals
                    } else {
                        &chunk.globals
                    };
                    if let Some(name) = globals.get(chunk.read_u32(offset + 4) as usize)
                        && let Some(&target) = functions.get(name.as_str())
                    {
                        pending.push(target);
                    }
                }
                Op::REF_FUNC => {
                    let target = chunk.read_u16(offset + 4) as usize;
                    assert!(target < chunks.len(), "function reference in {}", chunk.name);
                    pending.push(target);
                }
                _ => {},
            }
            let (_, next) = vybe_runtime::debug::disassemble_instruction(chunk, offset);
            assert!(next > offset && next <= chunk.code.len());
            offset = next;
        }
    }
    dumps.join("\n")
}

#[test]
fn realloc_null_and_size_guard_use_static_comparisons() {
    let path = Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../tests/c/doom_startup/realloc_null_static_guard.c");
    vybe_language_c::register();
    let chunks = vybe_compiler::projects::load(&path).unwrap().compile().unwrap();
    for name in ["allocate", "reject"] {
        let chunk = chunks.iter().find(|chunk| chunk.name == name).unwrap();
        let dump = vybe_runtime::debug::disassemble(chunk);
        assert!(!dump.contains("dyneq"), "dynamic realloc equality in {name}: {dump}");
        assert!(!dump.contains("__vybe_slot_"), "dynamic realloc size guard in {name}: {dump}");
        assert!(dump.contains("f64.le"), "missing static allocation guard: {dump}");
    }
    let bytes = vybe_platform_wasm::write_wasm_with_options(&chunks,
        vybe_platform_wasm::WasmWriteOptions {
            include_vybe_metadata: false,
            include_name_section: true,
        });
    validate_wasm_with_context(&bytes, &chunks);
    let mut vm = vybe_runtime::VM::new();
    vybe_runtime::init_all_registered(&mut vm, &vybe_runtime::capabilities::Capabilities::all());
    vm.run(chunks).unwrap();
    assert_eq!(vm.pending_exit_code, 0);
}

#[test]
fn realloc_integer_buffers_preserve_prefix_and_static_storage() {
    let path = Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../tests/c/doom_startup/realloc_integer_prefix.c");
    vybe_language_c::register();
    let chunks = vybe_compiler::projects::load(&path).unwrap().compile().unwrap();
    assert_no_dynamic_equality(&chunks);
    let bytes = vybe_platform_wasm::write_wasm_with_options(&chunks,
        vybe_platform_wasm::WasmWriteOptions {
            include_vybe_metadata: false,
            include_name_section: true,
        });
    validate_wasm_with_context(&bytes, &chunks);
    let mut vm = vybe_runtime::VM::new();
    vybe_runtime::init_all_registered(&mut vm, &vybe_runtime::capabilities::Capabilities::all());
    vm.run(chunks).unwrap();
    assert_eq!(vm.pending_exit_code, 0);
}

#[test]
fn computed_scalar_pointer_assignments_keep_linear_backing() {
    let path = Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../tests/c/doom_startup/linear_scalar_column_copy.c");
    let mut vm = vybe_runtime::VM::new();
    vybe_runtime::init_all_registered(
        &mut vm,
        &vybe_runtime::capabilities::Capabilities::all(),
    );
    vybe_language_c::register();
    let chunks = vybe_compiler::projects::load(&path).unwrap().compile().unwrap();
    let helper = chunks.iter().find(|chunk| chunk.name == "copy_column").unwrap();
    let dump = vybe_runtime::debug::disassemble(helper);
    assert!(helper.code.len() < 1000, "scalar column copy expanded: {dump}");
    assert!(dump.contains("i32.load16_s") && dump.contains("i32.store16"), "missing linear short copy: {dump}");
    assert!(!dump.contains("struct.new"), "linear column copy creates pointer wrappers: {dump}");
    assert!(!dump.contains("array.set"), "linear column copy uses managed storage: {dump}");
    assert!(!dump.contains("dyneq"), "linear column copy uses dynamic equality: {dump}");
    assert!(!dump.contains("__vybe_autoderef"), "integer loop condition uses reference checks: {dump}");
    assert!(!dump.contains("f64.add"), "linear copy addresses use floating arithmetic: {dump}");
    assert!(!dump.contains("f64.div"), "linear copy contains repeated scalar casts: {dump}");
    eprintln!("copy_column bytecode bytes: {}", helper.code.len());
    vm.run(chunks).unwrap_or_else(|error| panic!("{error}\n{dump}"));
    assert_eq!(vm.pending_exit_code, 0);
}

#[test]
#[ignore = "compiles the complete local Doom source"]
fn doom_load_menu_keeps_row_addresses() {
    let path = Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../tests/c/doom_startup/full_save_menu.c");
    vybe_language_c::register();
    let mut vm = vybe_runtime::VM::new();
    vybe_runtime::init_all_registered(&mut vm, &vybe_runtime::capabilities::Capabilities::all());
    let compiled = vybe_compiler::dynamic::RuntimeCompilerService::new(&mut vm)
        .compile_path(&path).unwrap();
    assert_no_dynamic_equality(&compiled.chunks);
    let bytes = vybe_platform_wasm::write_wasm_with_options(&compiled.chunks,
        vybe_platform_wasm::WasmWriteOptions {
            include_vybe_metadata: false,
            include_name_section: true,
        });
    validate_wasm_with_context(&bytes, &compiled.chunks);
    vybe_compiler::dynamic::RuntimeCompilerService::new(&mut vm)
        .run_compiled(compiled).unwrap();
    assert_eq!(vm.pending_exit_code, 0);
}

#[test]
#[ignore = "compiles the complete local Doom source"]
fn doom_full_program_uses_static_equality_and_valid_wasm() {
    let chunks = compile_and_validate_static_doom();
    println!("full Doom static-equality gate: {} chunks", chunks.len());
}

fn compile_and_validate_static_doom() -> Vec<vybe_runtime::Chunk> {
    let path = Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../examples/chocolate-doom/vybe_doom_core.c");
    vybe_language_c::register();
    let chunks = vybe_compiler::projects::load(&path).unwrap().compile().unwrap();
    assert_no_dynamic_equality(&chunks);
    let bytes = vybe_platform_wasm::write_wasm_with_options(&chunks,
        vybe_platform_wasm::WasmWriteOptions {
            include_vybe_metadata: false,
            include_name_section: true,
        });
    validate_wasm_with_context(&bytes, &chunks);
    if let Some(path) = std::env::var_os("C_DOOM_STATIC_WASM_PATH") {
        std::fs::write(path, &bytes).unwrap();
    }
    chunks
}

#[test]
#[ignore = "compiles the complete local Doom source"]
fn doom_melt_uses_linear_source_pixels() {
    let chunks = compile_and_validate_static_doom();
    let helper = chunks.iter().find(|chunk| chunk.name == "wipe_doMelt").unwrap();
    let dump = vybe_runtime::debug::disassemble(helper);
    assert!(helper.code.len() < 6000, "linear melt source expanded: {dump}");
    assert!(dump.contains("i32.load16_s"), "missing linear melt source: {dump}");
    assert!(!dump.contains("__c_scalar_read_ptr"), "melt source uses backing checks: {dump}");
    assert!(!dump.contains("dyneq"), "melt uses dynamic equality: {dump}");
    println!("wipe_doMelt bytecode bytes: {}", helper.code.len());
    for name in ["R_DrawSpan", "R_DrawColumn"] {
        let chunk = chunks.iter().find(|chunk| chunk.name == name).unwrap();
        let dump = vybe_runtime::debug::disassemble(chunk);
        println!("{name} bytecode bytes: {}, f64.div: {}, f64.add: {}",
            chunk.code.len(), dump.matches("f64.div").count(), dump.matches("f64.add").count());
    }
    for name in ["R_FindPlane", "R_CheckPlane", "T_Glow", "P_UpdateSpecials",
        "G_ReadDemoTiccmd", "wipe_doColorXForm", "M_BaseName", "M_DirName",
        "M_FileCaseExists", "SetVideoMode", "ToggleFullScreenKeyShortcut",
        "GetTypedChar", "M_NormalizeSlashes", "HUlib_addLineToSText", "M_FileExists"] {
        let chunk = chunks.iter().find(|chunk| chunk.name == name).unwrap();
        let dump = vybe_runtime::debug::disassemble(chunk);
        assert!(!dump.contains("dyneq"), "dynamic plane comparison in {name}: {dump}");
    }
    let dynamic_renderers: Vec<_> = chunks.iter()
        .filter(|chunk| chunk.name.starts_with("R_"))
        .filter_map(|chunk| {
            let count = vybe_runtime::debug::disassemble(chunk).matches("dyneq").count();
            (count > 0).then(|| format!("{}: {count}", chunk.name))
        })
        .collect();
    assert!(dynamic_renderers.is_empty(),
        "renderer dynamic equality remains: {}", dynamic_renderers.join(", "));
    let mut dynamic_game = chunks.iter().filter(|chunk| {
        !chunk.name.starts_with("__") && !chunk.name.starts_with('<')
    }).filter_map(|chunk| {
        let count = vybe_runtime::debug::disassemble(chunk).matches("dyneq").count();
        (count > 0).then(|| (count, chunk.name.clone()))
    }).collect::<Vec<_>>();
    dynamic_game.sort_by(|left, right| right.0.cmp(&left.0));
    println!("Doom dynamic equality audit: {} functions; largest: {:?}",
        dynamic_game.len(), &dynamic_game[..dynamic_game.len().min(20)]);
    assert!(dynamic_game.is_empty(), "Doom source uses dynamic equality: {dynamic_game:?}");
    let mut dynamic_helpers = chunks.iter().filter(|chunk| {
        chunk.name.starts_with("__") || chunk.name.starts_with('<')
    }).filter_map(|chunk| {
        let count = vybe_runtime::debug::disassemble(chunk).matches("dyneq").count();
        (count > 0).then(|| (count, chunk.name.clone()))
    }).collect::<Vec<_>>();
    dynamic_helpers.sort_by(|left, right| right.0.cmp(&left.0));
    println!("Doom generated dynamic equality audit: {dynamic_helpers:?}");
    assert!(dynamic_helpers.is_empty(),
        "Doom generated helpers use dynamic equality: {dynamic_helpers:?}");
}

#[test]
fn heap_record_array_fields_decay_to_their_backing() {
    let path = Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../tests/c/doom_startup/heap_record_name_comparison.c");
    let mut vm = vybe_runtime::VM::new();
    vybe_runtime::init_all_registered(
        &mut vm,
        &vybe_runtime::capabilities::Capabilities::all(),
    );
    vybe_language_c::register();
    let chunks = vybe_compiler::projects::load(&path).unwrap().compile().unwrap();
    let matches = chunks.iter().find(|chunk| chunk.name == "matches").unwrap();
    let dump = vybe_runtime::debug::disassemble(matches);
    assert!(matches.code.len() < 4000, "record array decay expanded: {dump}");
    assert!(!dump.contains("dyneq"), "dynamic byte comparison: {dump}");
    vm.run(chunks).unwrap_or_else(|error| panic!("{error}\n{dump}"));
    assert_eq!(vm.pending_exit_code, 0);
}

#[test]
#[ignore = "requires the local Freedoom 1 WAD"]
fn wad_texture_column_matches_native_pixels() {
    let path = Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../tests/c/doom_startup/wad_texture_composite_column.c");
    let mut vm = vybe_runtime::VM::new();
    vybe_runtime::init_all_registered(
        &mut vm,
        &vybe_runtime::capabilities::Capabilities::all(),
    );
    vybe_language_c::register();
    let chunks = vybe_compiler::projects::load(&path).unwrap().compile().unwrap();
    vm.run(chunks).unwrap();
    assert_eq!(vm.pending_exit_code, 0);
}

#[test]
fn unsigned_comparison_does_not_repeat_casts() {
    let path = Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../tests/c/doom_startup/unsigned_cast_idempotent.c");
    let mut vm = vybe_runtime::VM::new();
    vybe_runtime::init_all_registered(
        &mut vm,
        &vybe_runtime::capabilities::Capabilities::all(),
    );
    vybe_language_c::register();
    let chunks = vybe_compiler::projects::load(&path).unwrap().compile().unwrap();
    let helper = chunks.iter().find(|chunk| chunk.name == "outside").unwrap();
    let dump = vybe_runtime::debug::disassemble(helper);
    assert!(dump.matches("f64.div").count() <= 4, "repeated unsigned casts: {dump}");
    assert!(!dump.contains("dyneq"), "dynamic unsigned comparison: {dump}");
    vm.run(chunks).unwrap();
    assert_eq!(vm.pending_exit_code, 0);
}

#[test]
fn file_mode_flags_use_static_comparisons() {
    let path = Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../tests/c/doom_startup/fread_linear_cached_range.c");
    let mut vm = vybe_runtime::VM::new();
    vybe_runtime::init_all_registered(
        &mut vm,
        &vybe_runtime::capabilities::Capabilities::all(),
    );
    vybe_language_c::register();
    let chunks = vybe_compiler::projects::load(&path).unwrap().compile().unwrap();
    let helper = chunks.iter().find(|chunk| chunk.name == "__c_fopen_h").unwrap();
    let dump = vybe_runtime::debug::disassemble(helper);
    assert!(dump.matches("(__vybe_dyneq)").count() <= 8, "dynamic file-mode flags: {dump}");
    vm.run(chunks).unwrap();
    assert_eq!(vm.pending_exit_code, 0);
}

#[test]
fn character_reads_have_static_comparisons() {
    let path = Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../tests/c/doom_startup/static_character_reads.c");
    let mut vm = vybe_runtime::VM::new();
    vybe_runtime::init_all_registered(
        &mut vm,
        &vybe_runtime::capabilities::Capabilities::all(),
    );
    vybe_language_c::register();
    let chunks = vybe_compiler::projects::load(&path).unwrap().compile().unwrap();
    for name in ["valid_name", "slash", "is_option", "byte_marker", "byte_equal", "last_separator", "first_separator",
        "previous_separator", "previous_byte", "next_separator"] {
        let chunk = chunks.iter().find(|chunk| chunk.name == name).unwrap();
        let dump = vybe_runtime::debug::disassemble(chunk);
        assert!(!dump.contains("dyneq"), "{name}: {dump}");
        if matches!(name, "first_separator" | "last_separator") {
            assert!(!dump.contains("__vybe_slot_"), "dynamic character search: {dump}");
        }
    }
    vm.run(chunks).unwrap();
    assert_eq!(vm.pending_exit_code, 0);
}

#[test]
fn indexed_record_integer_reads_keep_declared_element_types() {
    let path = Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../tests/c/doom_startup/indexed_record_integer_reads.c");
    vybe_language_c::register();
    let chunks = vybe_compiler::projects::load(&path).unwrap().compile().unwrap();
    for name in ["same_score", "negative_score", "high_byte", "next_score"] {
        let chunk = chunks.iter().find(|chunk| chunk.name == name).unwrap();
        let dump = vybe_runtime::debug::disassemble(chunk);
        assert!(!dump.contains("dyneq"), "{name}: {dump}");
    }
    let mut vm = vybe_runtime::VM::new();
    vybe_runtime::init_all_registered(&mut vm, &vybe_runtime::capabilities::Capabilities::all());
    vm.run(chunks).unwrap();
    assert_eq!(vm.pending_exit_code, 0);
}

#[test]
fn errno_header_has_typed_writable_storage_and_distinct_codes() {
    let path = Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../tests/c/doom_startup/errno_header_values.c");
    vybe_language_c::register();
    let chunks = vybe_compiler::projects::load(&path).unwrap().compile().unwrap();
    for name in ["error_is_directory", "distinct_errors"] {
        let chunk = chunks.iter().find(|chunk| chunk.name == name).unwrap();
        let dump = vybe_runtime::debug::disassemble(chunk);
        assert!(!dump.contains("dyneq"), "{name}: {dump}");
    }
    let mut vm = vybe_runtime::VM::new();
    vybe_runtime::init_all_registered(&mut vm, &vybe_runtime::capabilities::Capabilities::all());
    vm.run(chunks).unwrap();
    assert_eq!(vm.pending_exit_code, 0);
}

#[test]
fn record_updates_preserve_the_result_type_and_local_storage() {
    let path = Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../tests/c/doom_startup/record_update_result.c");
    vybe_language_c::register();
    let chunks = vybe_compiler::projects::load(&path).unwrap().compile().unwrap();
    for name in ["advance_count", "old_count", "wrap_byte", "high_byte", "retreat_flags", "advance_once"] {
        let chunk = chunks.iter().find(|chunk| chunk.name == name).unwrap();
        let dump = vybe_runtime::debug::disassemble(chunk);
        assert!(!dump.contains("dyneq"), "{name}: {dump}");
        assert!(!dump.lines().any(|line| line.contains("global.")
            && line.contains("__c_member_old")), "shared update temporary: {dump}");
    }
    let mut vm = vybe_runtime::VM::new();
    vybe_runtime::init_all_registered(&mut vm, &vybe_runtime::capabilities::Capabilities::all());
    vm.run(chunks).unwrap();
    assert_eq!(vm.pending_exit_code, 0);
}

#[test]
fn scalar_dereferences_keep_static_comparisons() {
    let path = Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../tests/c/doom_startup/static_scalar_dereferences.c");
    let mut vm = vybe_runtime::VM::new();
    vybe_runtime::init_all_registered(
        &mut vm,
        &vybe_runtime::capabilities::Capabilities::all(),
    );
    vybe_language_c::register();
    let chunks = vybe_compiler::projects::load(&path).unwrap().compile().unwrap();
    for name in ["equal_words", "sentinel", "resize_changed",
        "signed_short_sentinel", "unsigned_short_sentinel"] {
        let chunk = chunks.iter().find(|chunk| chunk.name == name).unwrap();
        let dump = vybe_runtime::debug::disassemble(chunk);
        assert!(!dump.contains("dyneq"), "{name}: {dump}");
    }
    let bytes = vybe_platform_wasm::write_wasm_with_options(&chunks,
        vybe_platform_wasm::WasmWriteOptions {
            include_vybe_metadata: false, include_name_section: true,
        });
    validate_wasm_with_context(&bytes, &chunks);
    vm.run(chunks).unwrap();
    assert_eq!(vm.pending_exit_code, 0);
}

#[test]
fn ctype_preserves_scalar_read_casts() {
    let path = Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../tests/c/doom_startup/static_ctype_reads.c");
    let mut vm = vybe_runtime::VM::new();
    vybe_runtime::init_all_registered(
        &mut vm,
        &vybe_runtime::capabilities::Capabilities::all(),
    );
    vybe_language_c::register();
    let chunks = vybe_compiler::projects::load(&path).unwrap().compile().unwrap();
    for name in ["whitespace", "classified", "control"] {
        let chunk = chunks.iter().find(|chunk| chunk.name == name).unwrap();
        let dump = vybe_runtime::debug::disassemble(chunk);
        assert!(!dump.contains("dyneq"), "{name}: {dump}");
    }
    assert!(!chunks.iter().any(|chunk| chunk.name.starts_with("__libc_ctype_")),
        "ctype operations must emit inline, not as guest helper functions");
    vm.run(chunks).unwrap();
    assert_eq!(vm.pending_exit_code, 0);
}

#[test]
fn stdio_status_uses_static_integer_flags() {
    let path = Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../tests/c/doom_startup/static_file_status.c");
    let mut vm = vybe_runtime::VM::new();
    vybe_runtime::init_all_registered(
        &mut vm,
        &vybe_runtime::capabilities::Capabilities::all(),
    );
    vybe_language_c::register();
    let chunks = vybe_compiler::projects::load(&path).unwrap().compile().unwrap();
    for name in ["eof_set", "error_set", "same_status"] {
        let chunk = chunks.iter().find(|chunk| chunk.name == name).unwrap();
        let dump = vybe_runtime::debug::disassemble(chunk);
        assert!(!dump.contains("dyneq"), "{name}: {dump}");
    }
    vm.run(chunks).unwrap();
    assert_eq!(vm.pending_exit_code, 0);
}

#[test]
fn bitwise_results_keep_integer_lanes() {
    let path = Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../tests/c/doom_startup/static_bitwise_results.c");
    let mut vm = vybe_runtime::VM::new();
    vybe_runtime::init_all_registered(
        &mut vm,
        &vybe_runtime::capabilities::Capabilities::all(),
    );
    vybe_language_c::register();
    let chunks = vybe_compiler::projects::load(&path).unwrap().compile().unwrap();
    for name in ["masked", "text_mask", "signed_bits", "signed_shift", "high_bits"] {
        let chunk = chunks.iter().find(|chunk| chunk.name == name).unwrap();
        let dump = vybe_runtime::debug::disassemble(chunk);
        assert!(!dump.contains("dyneq"), "{name}: {dump}");
    }
    let wide = chunks.iter().find(|chunk| chunk.name == "wide_bits").unwrap();
    let dump = vybe_runtime::debug::disassemble(wide);
    assert!(dump.contains("i64.and"), "wide operand was narrowed: {dump}");
    vm.run(chunks).unwrap();
    assert_eq!(vm.pending_exit_code, 0);
}

#[test]
fn indexed_scalar_helper_uses_static_width_comparisons() {
    let path = Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../tests/c/doom_startup/hybrid_indexed_scalar_load.c");
    let mut vm = vybe_runtime::VM::new();
    vybe_runtime::init_all_registered(
        &mut vm,
        &vybe_runtime::capabilities::Capabilities::all(),
    );
    vybe_language_c::register();
    let chunks = vybe_compiler::projects::load(&path).unwrap().compile().unwrap();
    let helper = chunks.iter()
        .find(|chunk| chunk.name == "__c_hybrid_indexed_scalar_load")
        .expect("fixture must exercise the compatibility load helper");
    let dump = vybe_runtime::debug::disassemble(helper);
    assert!(!dump.contains("dyneq"), "dynamic width comparison: {dump}");
    assert!(!dump.contains("__vybe_slot_"), "dynamic signed scalar conversion: {dump}");
    assert!(dump.contains("i32.shl") && dump.contains("i32.shr_s"), "missing signed extension: {dump}");
    assert!(dump.contains("i32.eq") || dump.contains("f64.eq"), "missing static width comparison: {dump}");
    let reads = chunks.iter().find(|chunk| chunk.name == "check_fields").unwrap();
    let reads_dump = vybe_runtime::debug::disassemble(reads);
    assert!(reads_dump.contains("i32.shr_s"), "missing signed byte extension: {reads_dump}");
    vm.run(chunks).unwrap();
    assert_eq!(vm.pending_exit_code, 0);
}

#[test]
fn scalar_record_comparisons_are_static() {
    let path = Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../tests/c/doom_startup/static_record_scalar_comparisons.c");
    let mut vm = vybe_runtime::VM::new();
    vybe_runtime::init_all_registered(
        &mut vm,
        &vybe_runtime::capabilities::Capabilities::all(),
    );
    vybe_language_c::register();
    let chunks = vybe_compiler::projects::load(&path).unwrap().compile().unwrap();
    for name in [
        "equal_count", "equal_flags", "equal_offset", "equal_byte",
        "equal_arithmetic", "equal_masked", "switch_count", "switch_flags",
        "switch_fallthrough",
        "switch_kind", "switch_typedef_kind",
    ] {
        let chunk = chunks.iter().find(|chunk| chunk.name == name).unwrap();
        let dump = vybe_runtime::debug::disassemble(chunk);
        assert!(!dump.contains("dyneq"), "{name}: {dump}");
        assert!(dump.contains("i32.eq") || dump.contains("f64.eq"), "{name}: {dump}");
        if name == "equal_count" {
            assert!(!dump.contains("f64.trunc"), "unnecessary load cast: {dump}");
        }
        if name.starts_with("switch_") {
            for helper in ["dynlt", "dynle", "dynge"] {
                assert!(!dump.contains(helper), "dynamic switch index: {dump}");
            }
        }
    }
    let pixel = chunks.iter().find(|chunk| chunk.name == "table_pixel").unwrap();
    let dump = vybe_runtime::debug::disassemble(pixel);
    assert!(!dump.contains("f64.div"), "byte read uses floating remainder: {dump}");
    assert!(dump.contains("i32.and"), "byte read lost its integer lane: {dump}");
    vm.run(chunks).unwrap();
    assert_eq!(vm.pending_exit_code, 0);
}

#[test]
fn shadowed_pointer_keeps_its_own_storage_type() {
    let source = include_str!("../../../tests/c/doom_startup/linear_record_pointer_field.c");
    let module = vybe_language_c::parse(source).unwrap();
    let body = module.body.iter().find_map(|statement| match &statement.kind {
        vybe_ast::StmtKind::FunctionDecl { name, body, .. } if name == "set_state" => Some(body),
        _ => None,
    }).unwrap();
    let declaration = body.iter().find_map(|statement| match &statement.kind {
        vybe_ast::StmtKind::VarDecl { declarations, .. } => declarations.iter().find(|declaration| {
            matches!(&declaration.pattern, vybe_ast::BindingPattern::Ident(name) if name == "state")
        }),
        _ => None,
    }).unwrap();
    assert!(declaration.type_hint.is_none(), "{:?}", declaration.type_hint);
}

#[test]
fn record_places_preserve_values_and_backings() {
    let root = Path::new(env!("CARGO_MANIFEST_DIR")).join("../../tests/c/doom_startup");
    for fixture in [
        "embedded_record_pointer_copy.c",
        "record_pointer_snapshot_assignment.c",
        "linked_record_compound_field.c",
        "deh_readline_real.c",
        "record_pointer_memcpy.c",
        "keyboard_character_lookup.c",
        "linear_boolean_assignment.c",
        "linear_record_pointer_field.c",
        "sector_line_pointer_table.c",
        "enum_record_array.c",
        "sdl_keyboard_translation_table.c",
        "sdl_keyboard_event_roundtrip.c",
        "threaded_linear_string.c",
        "heap_thinker_function_pointer_equality.c",
    ] {
        let mut vm = vybe_runtime::VM::new();
        vybe_runtime::init_all_registered(
            &mut vm,
            &vybe_runtime::capabilities::Capabilities::all(),
        );
        vybe_language_c::register();
        let bundle = vybe_compiler::projects::load(&root.join(fixture))
            .unwrap_or_else(|error| panic!("{fixture}: {error}"));
        let chunks = bundle.compile().unwrap_or_else(|error| panic!("{fixture}: {error}"));
        if fixture == "heap_thinker_function_pointer_equality.c" {
            let main = chunks.iter().find(|chunk| chunk.name == "main").unwrap();
            assert!(main.code.len() < 128 * 1024,
                "nested record places expanded to {} bytes", main.code.len());
            for name in ["main", "update"] {
                let chunk = chunks.iter().find(|chunk| chunk.name == name).unwrap();
                assert!(!vybe_runtime::debug::disassemble(chunk).contains("dyneq"),
                    "dynamic equality in {name}");
            }
        }
        if let Err(error) = vm.run(chunks) {
            let location = error.call_stack.first().and_then(|frame| {
                vm.chunks.iter().find(|chunk| chunk.name == frame.chunk_name)
                    .and_then(|chunk| chunk.get_line(frame.offset.saturating_sub(1)))
            });
            panic!("{fixture}: source line {location:?}: {error}");
        }
        assert_eq!(vm.pending_exit_code, 0, "{fixture}");
    }
}
