fn run_fixture(name: &str) {
    vybe_language_c::register();
    let path = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../tests/c/doom_startup")
        .join(name);
    let chunks = vybe_compiler::projects::load(&path)
        .unwrap()
        .compile()
        .unwrap();
    if matches!(name, "sdl_texture_packed_word.c" | "sdl_member_packed_word.c"
        | "sdl_scalar_alias_word_view.c") {
        for chunk in &chunks {
            let dump = vybe_runtime::debug::disassemble(chunk);
            assert!(!dump.contains("dyneq"), "{}: dynamic equality", chunk.name);
            let imports = if chunk.imports.is_empty() {
                &chunks[0].imports
            } else {
                &chunk.imports
            };
            let mut offset = 0;
            while offset + 4 <= chunk.code.len() {
                let op = vybe_runtime::Op::decode(
                    chunk.read_u16(offset), chunk.read_u16(offset + 2)).unwrap();
                if op == vybe_runtime::Op::CALL {
                    if let Some(import) = imports.get(chunk.read_u16(offset + 4) as usize) {
                        assert!(!(import.module == "ecma:value"
                            && matches!(import.name.as_str(), "abstractEq" | "abstractNe")),
                            "{}@{offset}: dynamic equality import", chunk.name);
                    }
                }
                let (_, next) = vybe_runtime::debug::disassemble_instruction(chunk, offset);
                assert!(next > offset);
                offset = next;
            }
        }
    }
    let layout_dump = (name == "initialized_scalar_byte_layout.c").then(|| {
        vybe_runtime::debug::disassemble(chunks.iter().find(|chunk| chunk.name == "main").unwrap())
    });
    let bytes = vybe_platform_wasm::write_wasm_with_options(
        &chunks,
        vybe_platform_wasm::WasmWriteOptions {
            include_vybe_metadata: false,
            include_name_section: true,
        },
    );
    if let Err(error) = wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
        .validate_all(&bytes)
    {
        let mut names = std::collections::HashMap::new();
        for payload in wasmparser::Parser::new(0).parse_all(&bytes) {
            if let Ok(wasmparser::Payload::CustomSection(section)) = payload {
                if let wasmparser::KnownCustom::Name(reader) = section.as_known() {
                    for name in reader {
                        if let Ok(wasmparser::Name::Function(map)) = name {
                            for entry in map {
                                let entry = entry.unwrap();
                                names.insert(entry.index, entry.name.to_string());
                            }
                        }
                    }
                }
            }
        }
        let mut function = 0;
        for payload in wasmparser::Parser::new(0).parse_all(&bytes) {
            if let Ok(wasmparser::Payload::ImportSection(section)) = &payload {
                function += section.clone().into_imports().filter(|import|
                    import.as_ref().is_ok_and(|import| matches!(import.ty,
                        wasmparser::TypeRef::Func(_) | wasmparser::TypeRef::FuncExact(_))))
                    .count() as u32;
            }
            if let Ok(wasmparser::Payload::CodeSectionEntry(body)) = payload {
                if body.range().contains(&error.offset()) {
                    let nearby = body.get_operators_reader().unwrap()
                        .into_iter_with_offsets().filter_map(Result::ok)
                        .filter(|(_, offset)| *offset >= error.offset().saturating_sub(60)
                            && *offset <= error.offset() + 16)
                        .collect::<Vec<_>>();
                    eprintln!("{name}: invalid WASM context: {nearby:?}");
                    eprintln!("function {function}: {:?}, locals: {:?}", names.get(&function),
                        body.get_locals_reader().unwrap().into_iter().collect::<Vec<_>>());
                }
                function += 1;
            }
        }
        panic!("{name}: {error}");
    }
    let mut vm = vybe_runtime::VM::new();
    vybe_runtime::init_all_registered(&mut vm, &vybe_runtime::capabilities::Capabilities::all());
    vm.run(chunks).unwrap();
    assert_eq!(vm.pending_exit_code, 0);
    if let Some(dump) = layout_dump {
        let dynamic: Vec<_> = dump.lines().filter(|line| line.contains("dyneq")).collect();
        assert!(dynamic.is_empty(), "static byte layout uses dynamic equality: {dynamic:?}");
        assert!(dump.contains("i32.load"), "missing linear integer reads");
    }
}

#[test]
fn sdl_scalar_alias_sizes_preserve_typed_word_views() {
    run_fixture("sdl_scalar_alias_word_view.c");
}

#[test]
fn sdl_heap_pointer_member_retains_allocation_backing() {
    run_fixture("sdl_member_packed_word.c");
}

#[test]
fn sdl_texture_member_cast_reads_packed_words_after_lock_and_blit() {
    run_fixture("sdl_texture_packed_word.c");
}

#[test]
fn narrow_string_numeric_escapes_preserve_bytes_and_utf8() {
    run_fixture("narrow_string_escape_bytes.c");
}

#[test]
fn linear_short_pointer_offsets_preserve_declared_element_stride() {
    run_fixture("linear_short_pointer_offsets.c");
}

#[test]
fn hud_ammo_pointer_observes_player_updates() {
    run_fixture("hud_ammo_pointer_alias.c");
}

#[test]
fn pistol_compound_ammo_update_reaches_hud_pointer() {
    run_fixture("doom_pistol_ammo_update.c");
}

#[test]
fn projection_tables_preserve_unsigned_angle_wrap_and_fixed_scale() {
    run_fixture("doom_projection_tables.c");
}

#[test]
fn wall_ray_angles_preserve_all_eight_octants() {
    run_fixture("doom_wall_angles.c");
}

#[test]
fn clipper_fragments_preserve_requested_bounds_and_record_copies() {
    run_fixture("cliprange_sentinel_scan.c");
}

#[test]
fn initialized_integer_arrays_preserve_byte_aliasing() {
    run_fixture("initialized_scalar_byte_layout.c");
}

#[test]
fn grazing_ray_matches_native_geometry() {
    run_fixture("doom_grazing_ray.c");
}

#[test]
fn doom_demo_signed_byte_cast_evaluates_postincrement_once() {
    run_fixture("doom_demo_signed_postincrement.c");
}

#[test]
fn scalar_pointer_parameters_write_and_update_both_backings() {
    run_fixture("scalar_pointer_parameter_writes.c");
}

#[test]
fn file_writes_preserve_byte_backings_sizes_and_positions() {
    run_fixture("fwrite_backings.c");
}

#[test]
fn constant_array_segments_honor_alignment_and_reserve_heap() {
    use vybe_ast::{BindingPattern, ExprKind, Literal, StmtKind};
    let module = vybe_language_c::parse(
        "static const int lead[1] = {1};\n\
         _Alignas(64) static const int aligned[1] = {2};\n\
         int main(void) { return lead[0] + aligned[0] - 3; }",
    ).unwrap();
    let segments: Vec<_> = module.body.iter().filter_map(|statement| {
        if let StmtKind::DataSegment { offset: Some(offset), bytes, .. } = &statement.kind {
            if let ExprKind::Lit(Literal::Int(address)) = offset.kind {
                return Some((address, bytes.clone()));
            }
        }
        None
    }).collect();
    assert_eq!(segments, vec![(1024, 1i32.to_le_bytes().to_vec()),
                            (1088, 2i32.to_le_bytes().to_vec())]);
    let heap = module.body.iter().find_map(|statement| {
        let StmtKind::VarDecl { declarations, .. } = &statement.kind else { return None };
        declarations.iter().find_map(|declaration| {
            if matches!(&declaration.pattern, BindingPattern::Ident(name) if name == "__c_heap_ptr") {
                return declaration.init.as_ref();
            }
            None
        })
    }).unwrap();
    assert!(matches!(heap.kind, ExprKind::Lit(Literal::Int(1104))));
}
