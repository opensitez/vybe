//! Real emitter/Owner tests. These do not substitute for whole-frontend/Wasm gates.
use std::sync::Arc;
use vybe_compiler::primitives::helper_templates::{Context, Owner, Pool};
use vybe_runtime::{Chunk, Op, Value, registry::emitter_results::EmitRequest};

fn count(chunks: &[Chunk]) -> usize {
    chunks
        .iter()
        .filter(|c| c.name == "__dart_json_clean")
        .count()
}
fn setup() -> Arc<vybe_runtime::profile::LanguageProfile> {
    vybe_runtime::profile::parse_profile_shared(vybe_language_dart::profile_source()).unwrap()
}
fn owner() -> Owner {
    Owner::new(
        vybe_language_dart::emitter::json_adapter::HELPER_DEFINITIONS,
        Arc::new(Pool::default()),
    )
}
fn selected(owner: &mut Owner, context: &Context<'_>, chunks: &mut Vec<Chunk>, line: u32) -> bool {
    owner.emit(
        "dart.json_encode",
        vybe_language_dart::parse,
        vybe_language_dart::emitter::dispatch::dispatch,
        1,
        EmitRequest::Value,
        context,
        chunks,
        0,
        line,
    )
}
#[test]
fn repeated_encode_constructs_one_recursive_body_and_preserves_value() {
    let profile = setup();
    let directives = vybe_ast::Directives::default();
    let context = Context {
        profile: &profile,
        directives: &directives,
        receiver: Default::default(),
        own: false,
        prototype: false,
    };
    let mut runs = Vec::new();
    for shared in [false, true] {
        let mut chunks = vec![Chunk::new("<script>")];
        let mut owner = owner();
        for (i, n) in [1.0, 3.0].into_iter().enumerate() {
            chunks[0].emit_f64_const(n, 7);
            chunks[0].emit_array_new_fixed(0, 1, 7);
            if shared {
                assert!(selected(&mut owner, &context, &mut chunks, 7));
            } else {
                vybe_language_dart::emitter::json_adapter::emit_dart_json_encode(
                    &mut chunks,
                    0,
                    1,
                    7,
                );
            }
            if i == 0 {
                chunks[0].emit_op(Op::DROP, 7);
            }
        }
        chunks[0].emit_op(Op::RETURN, 7);
        assert_eq!(count(&chunks), if shared { 1 } else { 2 });
        let wasm = vybe_platform_wasm::write_wasm_with_options(
            &chunks,
            vybe_platform_wasm::WasmWriteOptions {
                include_vybe_metadata: false,
                include_name_section: true,
            },
        );
        wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
            .validate_all(&wasm)
            .unwrap();
        let mut vm = vybe_runtime::VM::new();
        vybe_compiler::primitives::platforms::register_platforms_all(&mut vm);
        runs.push(vm.run(chunks).unwrap());
    }
    assert_eq!(runs[0], runs[1]);
    assert!(matches!(&runs[0], Value::String(s) if s.as_ref() == "[3]"));
}
#[test]
fn owned_index_line_and_module_boundaries_are_authoritative() {
    let profile = setup();
    let directives = vybe_ast::Directives::default();
    let context = Context {
        profile: &profile,
        directives: &directives,
        receiver: Default::default(),
        own: false,
        prototype: false,
    };
    let mut chunks = vec![Chunk::new("caller"), Chunk::new("__dart_json_clean")];
    let mut first = owner();
    assert!(selected(&mut first, &context, &mut chunks, 17));
    assert_eq!(chunks.len(), 3, "guest spelling is never adopted");
    assert!(selected(&mut first, &context, &mut chunks, 17));
    assert_eq!(chunks.len(), 3);
    assert!(selected(&mut first, &context, &mut chunks, 18));
    assert_eq!(chunks.len(), 4, "source diagnostics require another body");
    assert!(selected(&mut owner(), &context, &mut chunks, 17));
    assert_eq!(
        chunks.len(),
        5,
        "another owner cannot adopt an existing body"
    );
}
#[test]
fn optional_callback_wrong_parser_and_arity_leave_original_route_untouched() {
    fn other(_: &str) -> Result<vybe_ast::Module, String> {
        Err("not Dart".into())
    }
    let profile = setup();
    let directives = vybe_ast::Directives::default();
    let context = Context {
        profile: &profile,
        directives: &directives,
        receiver: Default::default(),
        own: false,
        prototype: false,
    };
    for (parser, argc) in [
        (other as vybe_runtime::registry::ParseFn, 1),
        (vybe_language_dart::parse, 0),
        (vybe_language_dart::parse, 2),
        (vybe_language_dart::parse, 3),
    ] {
        let mut chunks = vec![Chunk::new("caller")];
        let before = format!("{chunks:?}");
        assert!(!owner().emit(
            "dart.json_encode",
            parser,
            vybe_language_dart::emitter::dispatch::dispatch,
            argc,
            EmitRequest::Value,
            &context,
            &mut chunks,
            0,
            7
        ));
        assert_eq!(format!("{chunks:?}"), before);
    }
    let mut chunks = vec![Chunk::new("caller")];
    let before = format!("{chunks:?}");
    assert!(!Owner::default().emit(
        "dart.json_encode",
        vybe_language_dart::parse,
        vybe_language_dart::emitter::dispatch::dispatch,
        1,
        EmitRequest::Value,
        &context,
        &mut chunks,
        0,
        7
    ));
    assert_eq!(format!("{chunks:?}"), before);
}

#[test]
fn actual_frontend_selected_compiler_avoids_second_build_without_changing_ast() {
    vybe_language_dart::register();
    let source =
        "import 'dart:convert'; void main() { print(jsonEncode([1, 2])); print(jsonEncode([3])); }";
    let module = vybe_language_dart::parse(source).unwrap();
    let before_ast = format!("{module:?}");
    let compile = |shared: bool| {
        let profile =
            vybe_compiler::profile::parse_profile(vybe_language_dart::profile_source()).unwrap();
        let compiler = vybe_compiler::Compiler::with_profile(profile);
        let compiler = if shared {
            compiler.with_helper_templates(
                vybe_language_dart::parse,
                vybe_language_dart::emitter::json_adapter::HELPER_DEFINITIONS,
                Arc::new(Pool::default()),
            )
        } else {
            compiler
        };
        compiler.compile(&module).unwrap()
    };
    let before = compile(false);
    let after = compile(true);
    assert_eq!(count(&before), 2);
    assert_eq!(count(&after), 1);
    assert!(
        after.iter().map(|c| c.code.len()).sum::<usize>()
            < before.iter().map(|c| c.code.len()).sum::<usize>()
    );
    assert_eq!(format!("{module:?}"), before_ast);
}
