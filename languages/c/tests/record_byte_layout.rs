use vybe_ast::{RecordByteLayout, StmtKind};

#[test]
fn pointer_offsets_retain_string_argument_type() {
    vybe_language_c::register();
    let path = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../tests/c/doom_startup/char_pointer_offset_string_argument.c");
    let chunks = vybe_compiler::projects::load(&path).unwrap().compile().unwrap();
    let mut vm = vybe_runtime::VM::new();
    vybe_runtime::init_all_registered(&mut vm, &vybe_runtime::capabilities::Capabilities::all());
    vm.run(chunks).unwrap();
    assert_eq!(vm.pending_exit_code, 0);
}

#[test]
fn mapped_record_checksum_reads_declared_fields() {
    vybe_language_c::register();
    let path = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../tests/c/doom_startup/mapped_record_byte_offsets.c");
    let chunks = vybe_compiler::projects::load(&path).unwrap().compile().unwrap();
    let mut vm = vybe_runtime::VM::new();
    vybe_runtime::init_all_registered(&mut vm, &vybe_runtime::capabilities::Capabilities::all());
    vm.run(chunks).unwrap();
    assert_eq!(vm.pending_exit_code, 0);
}

#[test]
fn mapped_record_checksum_wasm_is_valid_and_reloadable() {
    vybe_language_c::register();
    let path = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../tests/c/doom_startup/mapped_record_byte_offsets.c");
    let chunks = vybe_compiler::projects::load(&path).unwrap().compile().unwrap();
    let bytes = vybe_platform_wasm::write_wasm_with_options(&chunks,
        vybe_platform_wasm::WasmWriteOptions {
            include_vybe_metadata: false,
            include_name_section: true,
        });
    if let Err(error) = wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
        .validate_all(&bytes) {
        let mut body_index = 0;
        for payload in wasmparser::Parser::new(0).parse_all(&bytes) {
            if let wasmparser::Payload::CodeSectionEntry(body) = payload.unwrap() {
                if body.range().contains(&error.offset()) {
                    let mut reader = body.get_operators_reader().unwrap();
                    let mut previous = std::collections::VecDeque::new();
                    while !reader.eof() {
                        let offset = reader.original_position();
                        let operator = reader.read().unwrap();
                        previous.push_back(format!("{offset}: {operator:?}"));
                        if previous.len() > 12 { previous.pop_front(); }
                        if offset >= error.offset() {
                            panic!("WASM body {body_index}: {error}\n{}", previous.into_iter().collect::<Vec<_>>().join("\n"));
                        }
                    }
                }
                body_index += 1;
            }
        }
        panic!("{error}");
    }
    let chunks = vybe_platform_wasm::read_wasm(&bytes).unwrap();
    let mut vm = vybe_runtime::VM::new();
    vybe_runtime::init_all_registered(&mut vm, &vybe_runtime::capabilities::Capabilities::all());
    vm.run(chunks).unwrap();
    assert_eq!(vm.pending_exit_code, 0);
}

#[test]
fn reference_byte_coordinates_do_not_change_field_aliasing() {
    vybe_language_c::register();
    let path = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../tests/c/doom_startup/record_reference_byte_coordinates.c");
    let chunks = vybe_compiler::projects::load(&path)
        .unwrap()
        .compile()
        .unwrap();
    let mut vm = vybe_runtime::VM::new();
    vybe_runtime::init_all_registered(&mut vm, &vybe_runtime::capabilities::Capabilities::all());
    vm.run(chunks).unwrap();
    assert_eq!(vm.pending_exit_code, 0);
}

fn layout(source: &str, record: &str) -> Option<RecordByteLayout> {
    let module = vybe_language_c::parse(source).unwrap();
    module
        .body
        .iter()
        .find_map(|statement| match &statement.kind {
            StmtKind::StructDecl {
                name, semantics, ..
            } if name == record => Some(semantics.byte_layout.clone()),
            _ => None,
        })
        .unwrap_or_else(|| panic!("missing record {record}"))
}

#[test]
fn padded_records_publish_offsets_and_array_extents() {
    let layout = layout(
        "typedef struct { char tag; int number; short samples[3]; } Record;",
        "Record",
    )
    .unwrap();
    layout.validate().unwrap();
    assert_eq!((layout.bytes, layout.align), (16, 4));
    assert_eq!(
        (
            layout.field("tag").unwrap().offset,
            layout.field("tag").unwrap().bytes
        ),
        (0, 1)
    );
    assert_eq!(
        (
            layout.field("number").unwrap().offset,
            layout.field("number").unwrap().bytes
        ),
        (4, 4)
    );
    assert_eq!(
        (
            layout.field("samples").unwrap().offset,
            layout.field("samples").unwrap().bytes
        ),
        (8, 6)
    );
}

#[test]
fn union_members_share_offset_without_losing_extent() {
    let layout = layout("typedef union { char byte; int number; } Union;", "Union").unwrap();
    assert_eq!((layout.bytes, layout.align), (4, 4));
    assert_eq!(layout.field("byte").unwrap().offset, 0);
    assert_eq!(layout.field("number").unwrap().offset, 0);
}

#[test]
fn bitfields_do_not_publish_guessed_byte_coordinates() {
    assert!(layout(
        "typedef struct { unsigned int a:3; unsigned int b:5; } Bits;",
        "Bits"
    )
    .is_none());
}

#[test]
fn nested_callback_union_retains_padding_in_the_containing_record() {
    let source = "typedef enum { FIRST = 1, SECOND = 2 } Sprite;
        typedef void (*Callback)(void);
        typedef union { Callback call; void (*argument)(void *); } Action;
        typedef struct { Sprite sprite; int frame; int tics; Action action;
            int next; int misc1; int misc2; } State;";
    let action = layout(source, "Action").unwrap();
    let state = layout(source, "State").unwrap();
    let member = state.field("action").unwrap();
    assert_eq!(state.field("sprite").unwrap().bytes, 4);
    assert_eq!(member.bytes, action.bytes);
    assert_eq!(member.offset % action.align, 0);
    assert_eq!(
        state.field("next").unwrap().offset,
        member.offset + action.bytes
    );
    assert_eq!(
        state.field("misc2").unwrap().offset,
        member.offset + action.bytes + 8
    );
    state.validate().unwrap();
}

#[test]
fn shared_normalization_rejects_layout_for_an_undeclared_field() {
    vybe_language_c::register();
    let mut module = vybe_language_c::parse("typedef struct { int x; } Record;").unwrap();
    let declaration = module
        .body
        .iter_mut()
        .find_map(|statement| match &mut statement.kind {
            StmtKind::StructDecl {
                name, semantics, ..
            } if name == "Record" => Some(semantics),
            _ => None,
        })
        .unwrap();
    declaration.byte_layout.as_mut().unwrap().fields[0].name = "missing".into();
    let profile = vybe_compiler::profile::parse_profile(vybe_language_c::profile_source()).unwrap();
    let error = vybe_compiler::Compiler::with_profile(profile)
        .compile(&module)
        .unwrap_err();
    assert!(error.contains("undeclared field missing"), "{error}");
}
