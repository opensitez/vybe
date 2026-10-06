use vybe_compiler::primitives::{bundle, callable, csv, functions, globals, platforms};
use vybe_runtime::chunk::{LocalName, LocalOrigin};
use vybe_runtime::value::{Object, ObjectKind};
use vybe_runtime::{Chunk, Op, VM, Value};
#[path = "fixtures/csv_private_baseline.rs"]
mod baseline;
#[derive(Clone, Copy, Debug)]
enum Operation {
    Line,
    Document,
    Minimal,
    Whitespace,
}
const OPERATIONS: [Operation; 4] = [
    Operation::Line,
    Operation::Document,
    Operation::Minimal,
    Operation::Whitespace,
];
impl Operation {
    fn names(self) -> (&'static str, &'static str) {
        match self {
            Self::Line => ("__stdlib_csv_parse_line", "__vybe_csv_parse_line"),
            Self::Document => ("__stdlib_csv_parse_document", "__vybe_csv_parse_document"),
            Self::Minimal => ("__stdlib_csv_format_minimal", "__vybe_csv_format_minimal"),
            Self::Whitespace => (
                "__stdlib_csv_format_whitespace",
                "__vybe_csv_format_whitespace",
            ),
        }
    }
    fn input(self) -> Value {
        match self {
            Self::Line => string("a,\"b,c\",d"),
            Self::Document => string("a,b\r\nc,d\r\n"),
            Self::Minimal | Self::Whitespace => array(vec![string("has space"), string("x,y")]),
        }
    }
    fn expected(self) -> Value {
        match self {
            Self::Line => array(vec![string("a"), string("b,c"), string("d")]),
            Self::Document => array(vec![
                array(vec![string("a"), string("b")]),
                array(vec![string("c"), string("d")]),
            ]),
            Self::Minimal => string("has space,\"x,y\""),
            Self::Whitespace => string("\"has space\",\"x,y\""),
        }
    }
}
fn string(s: &str) -> Value {
    Value::String(s.into())
}
fn array(a: Vec<Value>) -> Value {
    Value::Object(vybe_runtime::heap::alloc(Object::new_array(a)))
}
fn same(a: &Value, b: &Value) {
    match (a, b) {
        (Value::Object(a), Value::Object(b)) => {
            let a = a.lock().unwrap();
            let b = b.lock().unwrap();
            let (ObjectKind::Array(a), ObjectKind::Array(b)) = (&a.kind, &b.kind) else {
                panic!("array results")
            };
            assert_eq!(a.len(), b.len());
            for (a, b) in a.iter().zip(b) {
                same(a, b)
            }
        }
        _ => assert_eq!(a, b),
    }
}
fn emit(operation: Operation, chunks: &mut [Chunk], current: usize, line: u32) {
    match operation {
        Operation::Line => csv::emit_parse_line(chunks, current, line),
        Operation::Document => csv::emit_parse_document(chunks, current, line),
        Operation::Minimal => {
            csv::emit_format_row(chunks, current, csv::FormatOptions::minimal(), line)
        }
        Operation::Whitespace => csv::emit_format_row(
            chunks,
            current,
            csv::FormatOptions::quote_whitespace(),
            line,
        ),
    }
}
fn frozen_factory(operation: Operation) -> Chunk {
    // Actual frozen algorithm in its original three-argument factory ABI.
    let mut c = Chunk::new(operation.names().0);
    c.arity = 3;
    c.local_count = 3;
    for slot in 0..3 {
        c.emit_op_u16(Op::LOCAL_GET, slot, 0)
    }
    let mut chunks = vec![c];
    match operation {
        Operation::Line => baseline::emit_parse_line(&mut chunks, 0, 0),
        Operation::Document => baseline::emit_parse_document(&mut chunks, 0, 0),
        Operation::Minimal => {
            baseline::emit_format_row(&mut chunks, 0, baseline::FormatOptions::minimal(), 0)
        }
        Operation::Whitespace => baseline::emit_format_row(
            &mut chunks,
            0,
            baseline::FormatOptions::quote_whitespace(),
            0,
        ),
    }
    chunks[0].emit_op(Op::RETURN, 0);
    chunks.pop().unwrap()
}
fn frozen_call(chunk: &mut Chunk, index: usize) {
    let slot = chunk.alloc_scratch(3);
    for arg in (0..3).rev() {
        chunk.emit_op_u16(Op::LOCAL_SET, slot + arg, 37)
    }
    functions::emit_ref_func(chunk, index, 0, 37);
    for arg in 0..3 {
        chunk.emit_op_u16(Op::LOCAL_GET, slot + arg, 37)
    }
    callable::emit_direct_invoke_chunk(chunk, 3, 37);
}
fn finalize(chunks: &mut Vec<Chunk>) {
    bundle::finalize_with_runtime_helpers(chunks);
    globals::declare_free_globals(chunks);
    globals::normalize_global_table(chunks)
}
fn program(old: bool, operation: Operation, calls: usize, alias: bool) -> Vec<Chunk> {
    let mut chunks = vec![Chunk::new("<csv-private>")];
    if old {
        chunks.push(frozen_factory(operation))
    }
    if alias {
        let index = chunks.len();
        let mut user = Chunk::new(operation.names().0);
        user.arity = 3;
        user.local_count = 3;
        user.local_names.push(LocalName::new(
            0,
            format!("<private-helper-body:{}>", operation.names().0),
            LocalOrigin::Source,
        ));
        user.emit_f64_const(42.0, 11);
        user.emit_op(Op::RETURN, 11);
        chunks.push(user);
        functions::emit_ref_func(&mut chunks[0], index, 0, 12);
        globals::emit_write(&mut chunks[0], operation.names().1, 12);
    }
    for n in 0..calls {
        for name in ["input", "delimiter", "enclosure"] {
            globals::emit_read(&mut chunks[0], name, 37)
        }
        if old {
            frozen_call(&mut chunks[0], 1)
        } else {
            emit(operation, &mut chunks, 0, 37)
        }
        if n + 1 != calls {
            chunks[0].emit_op(Op::DROP, 37)
        }
    }
    chunks[0].emit_op(Op::RETURN, 37);
    finalize(&mut chunks);
    chunks
}
fn runtime(input: Value, delimiter: Value, enclosure: Value) -> VM {
    let mut vm = VM::new();
    platforms::register_platforms_all(&mut vm);
    vm.set_global("input", input);
    vm.set_global("delimiter", delimiter);
    vm.set_global("enclosure", enclosure);
    vm
}
fn native_result(chunks: &[Chunk], label: &str) -> Result<(), wasmparser::BinaryReaderError> {
    let bytes = vybe_platform_wasm::write_wasm(chunks);
    let result = wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
        .validate_all(&bytes)
        .map(|_| ());
    if let Err(error) = &result {
        diagnose_native_failure(chunks, &bytes, error, label);
    }
    result
}
fn validate(chunks: &[Chunk]) {
    native_result(chunks, "current/frozen primitive fixture").unwrap();
}
fn diagnose_native_failure(
    chunks: &[Chunk],
    bytes: &[u8],
    error: &wasmparser::BinaryReaderError,
    label: &str,
) {
    let named = vybe_platform_wasm::write_wasm_with_options(
        chunks,
        vybe_platform_wasm::WasmWriteOptions {
            include_name_section: true,
            ..Default::default()
        },
    );
    let mut names = std::collections::HashMap::new();
    for payload in wasmparser::Parser::new(0).parse_all(&named) {
        if let wasmparser::Payload::CustomSection(section) = payload.unwrap() {
            if let wasmparser::KnownCustom::Name(reader) = section.as_known() {
                for name in reader {
                    if let wasmparser::Name::Function(map) = name.unwrap() {
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
    for payload in wasmparser::Parser::new(0).parse_all(bytes) {
        if let wasmparser::Payload::ImportSection(section) = payload.unwrap() {
            for import in section.into_imports() {
                if matches!(
                    import.unwrap().ty,
                    wasmparser::TypeRef::Func(_) | wasmparser::TypeRef::FuncExact(_)
                ) {
                    function += 1
                }
            }
        }
    }
    for payload in wasmparser::Parser::new(0).parse_all(bytes) {
        if let wasmparser::Payload::CodeSectionEntry(body) = payload.unwrap() {
            if body.range().contains(&error.offset()) {
                eprintln!(
                    "CSV native {label}: function={function} name={:?} range={:?} error={error}",
                    names.get(&function),
                    body.range()
                );
                let mut operators = body.get_operators_reader().unwrap();
                while !operators.eof() {
                    let (operator, offset) = operators.read_with_offset().unwrap();
                    if offset + 140 >= error.offset() && offset <= error.offset() + 40 {
                        eprintln!("  wasm@{offset}: {operator:?}");
                        if let wasmparser::Operator::Call { function_index } = operator {
                            eprintln!("    target={:?}", names.get(&function_index));
                        }
                    }
                }
            }
            function += 1;
        }
    }
}
fn validate_source_pair(chunks: &[Chunk], op: Operation, kind: u8) {
    let mut before = chunks.to_vec();
    let index = before
        .iter()
        .position(|c| {
            c.local_names.iter().any(|n| {
                n.origin == LocalOrigin::Compiler
                    && n.name == format!("<private-helper-body:{}>", op.names().0)
            })
        })
        .unwrap();
    let mut frozen = frozen_factory(op);
    frozen.local_names = before[index].local_names.clone();
    before[index] = frozen;
    let old = native_result(
        &before,
        &format!("frozen-factory real Source {op:?} kind={kind}"),
    );
    let new = native_result(
        chunks,
        &format!("private-factory real Source {op:?} kind={kind}"),
    );
    old.unwrap();
    new.unwrap();
}
fn evaluate(
    old: bool,
    operation: Operation,
    input: Value,
    delimiter: &str,
    enclosure: &str,
) -> Value {
    runtime(input, string(delimiter), string(enclosure))
        .run(program(old, operation, 1, false))
        .unwrap()
}
#[test]
fn private_csv_preserves_all_variants_aliases_and_source_origin_body_collisions() {
    for op in OPERATIONS {
        for old in [true, false] {
            let chunks = program(old, op, 1, true);
            validate(&chunks);
            if !old {
                assert!(
                    !chunks[0]
                        .global_inits
                        .iter()
                        .any(|g| g.name == op.names().1)
                );
                assert_eq!(
                    chunks
                        .iter()
                        .filter(|c| c
                            .local_names
                            .iter()
                            .any(|n| n.origin == LocalOrigin::Compiler
                                && n.name == format!("<private-helper-body:{}>", op.names().0)))
                        .count(),
                    1
                )
            }
            let mut vm = runtime(op.input(), string(","), string("\""));
            same(&vm.run(chunks).unwrap(), &op.expected());
            let alias = vm.global(op.names().1).unwrap().clone();
            assert_eq!(
                vm.invoke_callback(&alias, &[op.input(), string(","), string("\"")]),
                Value::F64(42.0)
            );
        }
    }
}
#[test]
fn frozen_current_factory_preserves_utf16_dialects_empty_and_field_conversions() {
    for op in OPERATIONS {
        validate(&program(true, op, 1, false));
        validate(&program(false, op, 1, false));
        let inputs = match op {
            Operation::Line | Operation::Document => [
                "",
                "a,",
                "\"\"",
                "a\r\nb\nc\r",
                "\"a\n😀\",é\n",
                "é😀世\0x",
                "a::b::",
                "a;\"b;c\";",
                "\"a\"\"b\",c",
            ]
            .into_iter()
            .map(string)
            .collect::<Vec<_>>(),
            _ => vec![
                array(vec![]),
                array(vec![string("")]),
                array(vec![string("😀"), string("é\0x"), string("a\r\nb")]),
                array(vec![
                    Value::I32(42),
                    Value::Bool(true),
                    Value::Null,
                    string(" a "),
                ]),
            ],
        };
        for input in inputs {
            for (delim, enc) in [
                (",", "\""),
                (";", "\""),
                ("::", ""),
                ("😀", "😀"),
                ("", "'"),
                ("\0", "\0"),
            ] {
                let a = evaluate(true, op, input.clone(), delim, enc);
                let b = evaluate(false, op, input.clone(), delim, enc);
                same(&a, &b)
            }
        }
    }
}
#[test]
fn private_csv_preserves_invalid_operand_errors_and_source_diagnostics() {
    for op in OPERATIONS {
        let programs = [program(true, op, 1, false), program(false, op, 1, false)];
        for input in [
            Value::Null,
            Value::Undefined,
            Value::I32(12),
            Value::Bool(false),
        ] {
            let normalize = |mut e: vybe_runtime::VMError| {
                for frame in &mut e.call_stack {
                    frame.offset = 0;
                }
                e.to_string()
            };
            let a = runtime(input.clone(), string(","), string("\""))
                .run(programs[0].clone())
                .map_err(normalize);
            let b = runtime(input, string(","), string("\""))
                .run(programs[1].clone())
                .map_err(normalize);
            match (a, b) {
                (Ok(a), Ok(b)) => same(&a, &b),
                (Err(a), Err(b)) => assert_eq!(a, b),
                other => panic!("{op:?} operand mismatch {other:?}"),
            }
        }
    }
}
#[test]
fn private_csv_exclusions_are_required_and_finalization_is_idempotent() {
    for op in OPERATIONS {
        let mut root = Chunk::new("<excluded-csv>");
        globals::emit_read(&mut root, "input", 37);
        root.emit_string_const(",", 37);
        root.emit_string_const("\"", 37);
        let mut chunks = vec![root];
        emit(op, &mut chunks, 0, 37);
        chunks[0].emit_op(Op::RETURN, 37);
        bundle::finalize_with_runtime_helpers_excluding(&mut chunks, &[op.names().0]);
        globals::declare_free_globals(&mut chunks);
        globals::normalize_global_table(&mut chunks);
        let before = chunks.iter().map(|c| c.code.clone()).collect::<Vec<_>>();
        bundle::finalize_with_runtime_helpers_excluding(&mut chunks, &[op.names().0]);
        assert_eq!(
            before,
            chunks.iter().map(|c| c.code.clone()).collect::<Vec<_>>()
        );
        assert_eq!(chunks.iter().filter(|c| c.name == op.names().0).count(), 1);
        assert!(
            !chunks[0]
                .global_inits
                .iter()
                .any(|g| g.name == op.names().1)
        );
        validate(&chunks);
        same(
            &runtime(op.input(), string(","), string("\""))
                .run(chunks)
                .unwrap(),
            &op.expected(),
        );
    }
}
fn placeholder() -> Chunk {
    let mut c = Chunk::new("unused-csv-padding");
    c.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, 0);
    c.emit_op(Op::RETURN, 0);
    c
}
fn i32_block(chunk: &mut Chunk) {
    chunk.emit_i32_const(1, 37);
    chunk.emit_if_i32(37);
    chunk.emit_i32_const(1, 37);
    chunk.emit_else(37);
    chunk.emit_i32_const(0, 37);
    chunk.emit_end(37);
    chunk.emit_op(Op::DROP, 37);
}
#[test]
fn csv_exact_u16_boundaries_preserve_input_slots_and_document_appended_i32_metadata() {
    for op in OPERATIONS {
        for initial in [65535, 65536] {
            let mut c = Chunk::new("<csv-boundary>");
            c.local_count = 7;
            i32_block(&mut c);
            let mut chunks = vec![c];
            for n in 0..2 {
                globals::emit_read(&mut chunks[0], "input", 37);
                chunks[0].emit_string_const(",", 37);
                chunks[0].emit_string_const("\"", 37);
                emit(op, &mut chunks, 0, 37);
                if n == 0 {
                    chunks[0].emit_op(Op::DROP, 37)
                }
            }
            i32_block(&mut chunks[0]);
            chunks[0].emit_op(Op::RETURN, 37);
            let indirect = chunks[0].code.len();
            chunks[0].emit_op(Op::CALL_INDIRECT, 37);
            for byte in [0, 0, 1] {
                chunks[0].emit(byte, 37)
            }
            chunks[0]
                .call_indirect_sigs
                .insert(indirect, "->externref".into());
            chunks[0]
                .call_indirect_canon
                .insert(indirect, "csv:original-callback-type".into());
            chunks[0].emit_op(Op::DROP, 37);
            chunks[0].user_code_offset = Some(indirect as u32);
            chunks.resize(initial, placeholder());
            finalize(&mut chunks);
            let fits = initial == 65535;
            assert_eq!(chunks.len(), initial + usize::from(fits));
            assert_eq!(
                chunks[0].block_i32_results.len(),
                2 + if !fits && matches!(op, Operation::Document) {
                    2
                } else {
                    0
                }
            );
            for &key in &chunks[0].block_i32_results {
                assert_eq!(
                    Op::decode(chunks[0].read_u16(key), chunks[0].read_u16(key + 2)),
                    Some(Op::IF)
                );
                assert_eq!(&chunks[0].code[key + 4..key + 6], &[0, 1]);
            }
            assert_eq!(chunks[0].code.len(), chunks[0].lines.len());
            let indirect = chunks[0].user_code_offset.unwrap() as usize;
            assert_eq!(
                Op::decode(
                    chunks[0].read_u16(indirect),
                    chunks[0].read_u16(indirect + 2)
                ),
                Some(Op::CALL_INDIRECT)
            );
            assert_eq!(chunks[0].call_indirect_sigs[&indirect], "->externref");
            assert_eq!(
                chunks[0].call_indirect_canon[&indirect],
                "csv:original-callback-type"
            );
            let before = chunks[0].code.clone();
            bundle::finalize_with_runtime_helpers(&mut chunks);
            assert_eq!(chunks[0].code, before);
            validate(&chunks);
            same(
                &runtime(op.input(), string(","), string("\""))
                    .run(chunks)
                    .unwrap(),
                &op.expected(),
            );
        }
    }
}
#[test]
fn oversized_document_i32_metadata_coexists_with_real_forward_generator_handlers() {
    use vybe_compiler::primitives::generators;
    use vybe_runtime::chunk::StackSwitchHandler;
    let mut root = Chunk::new("<document-generator>");
    root.emit_string_const("a,b\n", 7);
    root.emit_string_const(",", 7);
    root.emit_string_const("\"", 7);
    let mut chunks = vec![root];
    csv::emit_parse_document(&mut chunks, 0, 7);
    chunks[0].emit_op(Op::DROP, 7);
    functions::emit_ref_func(&mut chunks[0], 1, 0, 9);
    chunks[0].emit_op(Op::CONT_NEW, 9);
    generators::emit_next(&mut chunks[0], 9);
    chunks[0].emit_op(Op::DROP, 9);
    chunks[0].emit_op(Op::RETURN, 9);
    let (&resume, handlers) = chunks[0].stack_switch_handlers.iter().next().unwrap();
    let target = handlers[0].label_index;
    chunks[0].user_code_offset = Some(resume as u32);
    let structural_resume = chunks[0].code.len();
    generators::emit_resume(&mut chunks[0], 10);
    chunks[0].stack_switch_handlers.insert(
        structural_resume,
        vec![StackSwitchHandler {
            kind: 0,
            tag_index: 0,
            label_index: 0,
        }],
    );
    chunks[0].emit_op(Op::RETURN, 10);
    let original_len = chunks[0].code.len();
    let mut generator = functions::create_function_chunk("<csv-yielding-body>", 0);
    generator.emit_f64_const(17.0, 11);
    generators::emit_suspend(&mut generator, 11);
    generator.emit_op(Op::DROP, 11);
    generator.emit_f64_const(19.0, 11);
    generator.emit_op(Op::RETURN, 11);
    chunks.push(generator);
    chunks.resize(65536, placeholder());
    finalize(&mut chunks);
    let delta = chunks[0].code.len() - original_len;
    assert!(delta > 0);
    assert_eq!(chunks[0].user_code_offset, Some((resume + delta) as u32));
    assert_eq!(
        chunks[0].stack_switch_handlers[&(resume + delta)][0].label_index,
        target + delta as u32
    );
    assert_eq!(
        chunks[0].stack_switch_handlers[&(structural_resume + delta)][0].label_index,
        0
    );
    assert_eq!(chunks[0].block_i32_results.len(), 1);
    let &predicate = chunks[0].block_i32_results.iter().next().unwrap();
    assert_eq!(
        Op::decode(
            chunks[0].read_u16(predicate),
            chunks[0].read_u16(predicate + 2)
        ),
        Some(Op::IF)
    );
    assert_eq!(&chunks[0].code[predicate + 4..predicate + 6], &[0, 1]);
    // This is the existing compiler's VM-absolute generator contract, already
    // known to be native-blocked. Full Wasm validation is covered separately
    // by every normal and exact-boundary CSV allocation/call fixture above.
    assert_eq!(
        runtime(Value::Null, string(","), string("\""))
            .run(chunks)
            .unwrap(),
        Value::F64(17.0)
    );
}
fn load(vm: &mut VM, chunks: Vec<Chunk>, linked: bool) -> Value {
    if linked {
        let imports = chunks[0]
            .imports
            .iter()
            .map(|i| vm.resolve_import_target(&i.module, &i.name).unwrap())
            .collect();
        vm.run_linked(chunks, imports).unwrap()
    } else {
        vm.run(chunks).unwrap()
    }
}
#[test]
fn retained_csv_callbacks_survive_cumulative_large_loads_and_later_alias_writes() {
    for op in OPERATIONS {
        let mut root = Chunk::new("<retain-csv>");
        let mut caller = Chunk::new("source-csv");
        caller.arity = 3;
        caller.local_count = 3;
        for slot in 0..3 {
            caller.local_names.push(LocalName::new(
                slot,
                format!("arg{slot}"),
                LocalOrigin::Source,
            ));
            caller.emit_op_u16(Op::LOCAL_GET, slot, 37)
        }
        functions::emit_ref_func(&mut root, 1, 0, 1);
        root.emit_op(Op::RETURN, 1);
        let mut chunks = vec![root, caller];
        emit(op, &mut chunks, 1, 37);
        chunks[1].emit_op(Op::RETURN, 37);
        finalize(&mut chunks);
        validate(&chunks);
        for linked in [false, true] {
            let mut vm = runtime(op.input(), string(","), string("\""));
            load(&mut vm, vec![placeholder(); 65536], linked);
            let retained = load(&mut vm, chunks.clone(), linked);
            let args = [op.input(), string(","), string("\"")];
            same(&vm.invoke_callback(&retained, &args), &op.expected());
            load(&mut vm, program(false, op, 1, true), linked);
            same(&vm.invoke_callback(&retained, &args), &op.expected());
            assert_eq!(
                vm.invoke_callback(&vm.global(op.names().1).unwrap().clone(), &args),
                Value::F64(42.0)
            );
        }
    }
}
fn python_program(op: Operation, kind: u8, padding: usize) -> Vec<Chunk> {
    static REGISTER: std::sync::Once = std::sync::Once::new();
    REGISTER.call_once(vybe_language_python::register);
    let (hook, args) = match op {
        Operation::Line => ("__py_csv_parse_line", "'a,\"b,c\",d', ',', '\"'"),
        Operation::Minimal => ("__py_csv_format_row", "['has space','x,y'], ',', '\"'"),
        _ => unreachable!("only production-declared Python CSV hooks"),
    };
    let call = format!("{hook}({args})");
    let call = match kind {
        0 => format!("answer={call}\n"),
        1 => format!("def calculate():\n    return {call}\nanswer=calculate()\n"),
        _ => format!("calculate=lambda: {call}\nanswer=calculate()\n"),
    };
    let source = format!(
        "def {}(content, delimiter, enclosure):\n    return 42.0\n{} = {}\nalias_answer = {}('x', ',', '\"')\n{call}",
        op.names().0,
        op.names().1,
        op.names().0,
        op.names().1
    );
    let mut module = vybe_language_python::parse(&source).unwrap();
    {
        // Eight real Source wrappers keep each emitted body within the native
        // validator's function-size limit. Include the empty wrappers in the
        // unpadded compile too, so only lambda chunks change the helper prefix.
        assert!(padding <= 8 * 8192);
        let mut prefix = Vec::with_capacity(8 + module.body.len());
        for wrapper in 0..8 {
            let count = padding / 8 + usize::from(wrapper < padding % 8);
            assert!(count <= 8192);
            let mut body = Vec::with_capacity(count);
            for _ in 0..count {
                body.push(vybe_ast::Statement::new(vybe_ast::StmtKind::Expr(
                    vybe_ast::Expression::new(vybe_ast::ExprKind::Lambda {
                        params: vec![],
                        body: vybe_ast::LambdaBody::Expr(Box::new(vybe_ast::Expression::new(
                            vybe_ast::ExprKind::Lit(vybe_ast::Literal::Null),
                        ))),
                        is_async: false,
                        captures: vec![],
                    }),
                )));
            }
            prefix.push(vybe_ast::Statement::new(vybe_ast::StmtKind::FunctionDecl {
                name: format!("csv_padding_wrapper_{wrapper}"),
                params: vec![],
                return_type: None,
                body,
                modifiers: Default::default(),
                handles: vec![],
                is_async: false,
                is_generator: false,
                is_sub: false,
            }));
        }
        prefix.extend(module.body);
        module.body = prefix;
    }
    module
        .body
        .push(vybe_ast::Statement::new(vybe_ast::StmtKind::Return(Some(
            vybe_ast::Expression::ident("answer"),
        ))));
    let profile =
        vybe_compiler::profile::parse_profile(vybe_language_python::profile_source()).unwrap();
    vybe_compiler::Compiler::with_profile(profile)
        .compile(&module)
        .unwrap()
}
#[test]
fn real_python_csv_hooks_root_function_lambda_preserve_source_aliases_and_native_behavior() {
    for op in [Operation::Line, Operation::Minimal] {
        for kind in 0..3 {
            let chunks = python_program(op, kind, 0);
            validate_source_pair(&chunks, op, kind);
            let mut vm = runtime(Value::Null, string(","), string("\""));
            same(&vm.run(chunks).unwrap(), &op.expected());
            assert_eq!(vm.global("alias_answer"), Some(&Value::F64(42.0)));
        }
    }
}
#[test]
fn oversized_real_source_csv_root_function_lambda_preserve_aliases_and_native_fallback() {
    for op in [Operation::Line, Operation::Minimal] {
        for kind in 0..3 {
            let unpadded = python_program(op, kind, 0);
            let helper_index = unpadded
                .iter()
                .position(|c| {
                    c.local_names.iter().any(|n| {
                        n.origin == LocalOrigin::Compiler
                            && n.name == format!("<private-helper-body:{}>", op.names().0)
                    })
                })
                .unwrap();
            let fixed_prefix = unpadded[..helper_index]
                .iter()
                .filter(|c| {
                    !c.local_names.iter().any(|n| {
                        n.origin == LocalOrigin::Compiler
                            && n.name == format!("<private-helper-body:{}>", op.names().0)
                    })
                })
                .count();
            let padding = 65536 - fixed_prefix;
            eprintln!(
                "CSV oversized Source {op:?} kind={kind}: unpadded_total={} private_index={helper_index} nonprivate_prefix={fixed_prefix} lambda_padding={padding}",
                unpadded.len()
            );
            let chunks = python_program(op, kind, padding);
            for (index, chunk) in chunks.iter().enumerate() {
                if chunk.name == "<lambda>" || chunk.name == "<lambda_factory>" {
                    assert!(index <= u16::MAX as usize, "Source lambda REF_FUNC must fit");
                }
            }
            for wrapper in 0..8 {
                let name = format!("csv_padding_wrapper_{wrapper}");
                let index = chunks.iter().position(|chunk| chunk.name == name)
                    .expect("actual Source padding wrapper");
                assert!(index <= u16::MAX as usize, "Source wrapper REF_FUNC must fit");
            }
            // Every discarded zero-capture padding expression contributes one
            // <lambda> chunk. All non-private chunks must survive; finalization
            // can omit several private bodies once the module crosses u16,
            // including helpers requested independently by the Python profile.
            let inventory = |module: &[Chunk]| {
                let mut counts = std::collections::BTreeMap::new();
                for chunk in module {
                    let marker = chunk.local_names.iter().find_map(|local| {
                        (local.origin == LocalOrigin::Compiler
                            && local.name.starts_with("<private-helper-body:"))
                        .then(|| local.name.clone())
                    });
                    *counts.entry((chunk.name.clone(), marker)).or_insert(0usize) += 1;
                }
                counts
            };
            let before = inventory(&unpadded);
            let mut after = inventory(&chunks);
            let lambda_key = ("<lambda>".to_string(), None);
            let lambda_count = after.get_mut(&lambda_key).expect("actual padding lambdas");
            *lambda_count = lambda_count.checked_sub(padding).expect("all padding lambdas emitted");
            if *lambda_count == 0 {
                after.remove(&lambda_key);
            }
            let mut removed_private_bodies = 0;
            for (key, count) in &before {
                let remaining = after.remove(key).unwrap_or(0);
                assert!(remaining <= *count, "unexpected duplicate Source/helper chunk: {key:?}");
                if remaining != *count {
                    assert!(key.1.is_some(), "fallback removed a non-private chunk: {key:?}");
                    eprintln!("CSV oversized removed private body {key:?}: {}", count - remaining);
                    removed_private_bodies += count - remaining;
                }
            }
            assert!(after.is_empty(), "unexpected non-padding chunks: {after:?}");
            assert!(removed_private_bodies > 0);
            assert_eq!(chunks.len(), unpadded.len() - removed_private_bodies + padding);
            assert!(!chunks.iter().any(|c| {
                c.local_names.iter().any(|n| {
                    n.origin == LocalOrigin::Compiler
                        && n.name == format!("<private-helper-body:{}>", op.names().0)
                })
            }));
            validate(&chunks);
            let mut vm = runtime(Value::Null, string(","), string("\""));
            same(&vm.run(chunks).unwrap(), &op.expected());
            assert_eq!(vm.global("alias_answer"), Some(&Value::F64(42.0)));
        }
    }
}
fn benchmark_program(old: bool, op: Operation, calls: usize, retained: bool) -> Vec<Chunk> {
    use vybe_runtime::chunk::{ConstExpr, GlobalInit};
    let mut chunks = vec![Chunk::new("<csv-public-private-benchmark>")];
    let caller = if retained {
        chunks.push(Chunk::new("<retained-csv-caller>"));
        1
    } else {
        0
    };
    if old {
        let helper = chunks.len();
        chunks.push(frozen_factory(op));
        chunks[0].global_inits.push(GlobalInit {
            name: op.names().1.into(),
            init: ConstExpr::RefFunc(helper),
        });
    }
    for n in 0..calls {
        for name in ["input", "delimiter", "enclosure"] {
            globals::emit_read(&mut chunks[caller], name, 37)
        }
        if old {
            baseline::emit_pre_migration_public_call(&mut chunks[caller], op.names().1, 3, 37)
        } else {
            emit(op, &mut chunks, caller, 37)
        }
        if n + 1 != calls {
            chunks[caller].emit_op(Op::DROP, 37)
        }
    }
    chunks[caller].emit_op(Op::RETURN, 37);
    if retained {
        functions::emit_ref_func(&mut chunks[0], caller, 0, 1);
        chunks[0].emit_op(Op::RETURN, 1);
    }
    finalize(&mut chunks);
    chunks
}
fn instruction_count(c: &Chunk) -> usize {
    let mut offset = 0;
    let mut count = 0;
    while offset < c.code.len() {
        let op = Op::decode(c.read_u16(offset), c.read_u16(offset + 2)).unwrap();
        offset += 4 + op.operand_format().size_in(&c.code, offset + 4);
        count += 1;
    }
    assert_eq!(offset, c.code.len());
    count
}
#[test]
fn pre_migration_public_caller_emission_audit_and_private_factory_counts() {
    for op in OPERATIONS {
        let programs = [
            benchmark_program(true, op, 64, false),
            benchmark_program(false, op, 64, false),
        ];
        for (which, chunks) in programs.iter().enumerate() {
            validate(chunks);
            let helper = chunks.iter().find(|c| c.name == op.names().0).unwrap();
            eprintln!(
                "CSV {op:?}64 {} caller_bytes={} caller_ops={} caller_slots={} helper_bytes={} helper_ops={} helper_slots={} globals={} wasm_bytes={}",
                if which == 0 {
                    "frozen-public"
                } else {
                    "private"
                },
                chunks[0].code.len(),
                instruction_count(&chunks[0]),
                chunks[0].local_count,
                helper.code.len(),
                instruction_count(helper),
                helper.local_count,
                chunks[0].global_inits.len(),
                vybe_platform_wasm::write_wasm(chunks).len()
            );
            same(
                &runtime(op.input(), string(","), string("\""))
                    .run(chunks.clone())
                    .unwrap(),
                &op.expected(),
            );
        }
    }
}
#[test]
#[ignore = "interleaved opt-level-0 exact frozen PUBLIC caller/factory; cold and loaded phases"]
fn debug_private_csv_frozen_shared_factory_comparison() {
    for op in OPERATIONS {
        for retained in [false, true] {
            let programs = [
                benchmark_program(true, op, 64, retained),
                benchmark_program(false, op, 64, retained),
            ];
            let mut cold = [vec![], vec![]];
            let mut invoked = [vec![], vec![]];
            for round in 0..7 {
                for which in [round % 2, 1 - round % 2] {
                    let mut vm = runtime(op.input(), string(","), string("\""));
                    let chunks = programs[which].clone();
                    let start = std::time::Instant::now();
                    let result = vm.run(chunks).unwrap();
                    cold[which].push(start.elapsed());
                    if retained {
                        for _ in 0..8 {
                            let start = std::time::Instant::now();
                            let answer = vm.invoke_callback(&result, &[]);
                            invoked[which].push(start.elapsed());
                            same(&answer, &op.expected());
                        }
                    } else {
                        same(&result, &op.expected());
                    }
                }
            }
            for samples in &mut cold {
                samples.sort();
            }
            eprintln!(
                "CSV {op:?}64 {} cold/load frozen_public={}us private={}us speedup={:.2}x",
                if retained {
                    "retained-caller"
                } else {
                    "run-caller"
                },
                cold[0][3].as_micros(),
                cold[1][3].as_micros(),
                cold[0][3].as_secs_f64() / cold[1][3].as_secs_f64()
            );
            if retained {
                for samples in &mut invoked {
                    samples.sort();
                }
                let median = invoked[0].len() / 2;
                eprintln!(
                    "CSV {op:?}64 loadedinvoke frozen_public={}us private={}us speedup={:.2}x",
                    invoked[0][median].as_micros(),
                    invoked[1][median].as_micros(),
                    invoked[0][median].as_secs_f64() / invoked[1][median].as_secs_f64()
                );
            }
        }
    }
}
