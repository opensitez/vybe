use std::path::Path;

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
        let code = &helper.code;
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
            assert_eq!(narrowings, 3, "nested i32 address is narrowed again: {dump}");
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
    vm.run(chunks.clone()).unwrap();
    assert_eq!(vm.pending_exit_code, 0);
    assert!(!dump.contains("f64.div"), "redundant unsigned binding conversion: {dump}");
    assert!(!dump.contains("dyneq"), "dynamic span equality: {dump}");
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
        "ascii_text"] {
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
    let reloaded = vybe_platform_wasm::read_wasm(&wasm).unwrap();
    vm.run(reloaded).unwrap();
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
    for name in ["full", "before"] {
        let chunk = chunks.iter().find(|chunk| chunk.name == name).unwrap();
        let dump = vybe_runtime::debug::disassemble(chunk);
        assert!(!dump.contains("dyneq"), "dynamic pointer distance: {dump}");
    }
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
fn doom_melt_uses_linear_source_pixels() {
    let path = Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../examples/chocolate-doom/vybe_doom_core.c");
    vybe_language_c::register();
    let chunks = vybe_compiler::projects::load(&path).unwrap().compile().unwrap();
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
    for name in ["valid_name", "slash", "is_option", "byte_marker", "byte_equal", "last_separator",
        "previous_separator", "previous_byte", "next_separator"] {
        let chunk = chunks.iter().find(|chunk| chunk.name == name).unwrap();
        let dump = vybe_runtime::debug::disassemble(chunk);
        assert!(!dump.contains("dyneq"), "{name}: {dump}");
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
    for name in ["equal_words", "sentinel", "resize_changed"] {
        let chunk = chunks.iter().find(|chunk| chunk.name == name).unwrap();
        let dump = vybe_runtime::debug::disassemble(chunk);
        assert!(!dump.contains("dyneq"), "{name}: {dump}");
    }
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
