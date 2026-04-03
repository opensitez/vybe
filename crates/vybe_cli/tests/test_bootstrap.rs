/// Test the bootstrap approach: compile all → link → bootstrap chunk → single vm.run()
/// NOTE: Bootstrap requires unified import table (Linker needs to remap call_import indices).
/// Currently only works when all modules share the same import list (same compiler).
/// Full multi-language bootstrap needs Linker import unification — tracked for future.

use std::rc::Rc;
use std::cell::RefCell;
use vybe_bytecode::{VM, Value, HostContext, ImportTarget, Op, Chunk};

fn setup_vm() -> (VM, Rc<RefCell<Vec<String>>>) {
    let mut vm = VM::new();
    let output: Rc<RefCell<Vec<String>>> = Rc::new(RefCell::new(Vec::new()));
    let out = output.clone();
    vybe_host::register_all(&mut vm);
    vm.register_host_fn("wasi:cli", "log", Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
        let parts: Vec<String> = args.iter().map(|v| format!("{v}")).collect();
        out.borrow_mut().push(parts.join(" "));
        Value::Null
    }));
    vybe_host::setup_namespaces(&mut vm);
    (vm, output)
}

/// Build a bootstrap chunk that calls each component's script chunk in order.
/// Returns (bootstrap + all component chunks) with adjusted indices.
fn build_bootstrap(
    link_result: &vybe_bytecode::component::LinkResult,
    component_count: usize,
    entry_idx: Option<usize>,
) -> Vec<Chunk> {
    let mut bootstrap = Chunk::new("<bootstrap>");
    let line = 0u32;

    // Call library script chunks first (non-entry)
    for i in 0..component_count {
        if Some(i) == entry_idx { continue; }
        // script chunk is at component_offsets[i] + 1 (shifted by bootstrap)
        let script_idx = (link_result.component_offsets[i] + 1) as u16;
        bootstrap.emit_op_u16(Op::ref_func, script_idx, line);
        bootstrap.emit(0, line); // 0 upvalues
        bootstrap.emit_op_u8(Op::call_ref, 0, line);
        bootstrap.emit_op(Op::drop, line);
    }

    // Call entry script chunk last
    if let Some(ei) = entry_idx {
        let script_idx = (link_result.component_offsets[ei] + 1) as u16;
        bootstrap.emit_op_u16(Op::ref_func, script_idx, line);
        bootstrap.emit(0, line);
        bootstrap.emit_op_u8(Op::call_ref, 0, line);
        bootstrap.emit_op(Op::drop, line);
    }

    bootstrap.emit_op(Op::halt, line);
    bootstrap.local_count = 16;

    // Copy the unified import table from chunk 0 (set by Linker import unification).
    // The Linker already remapped all call_import indices to match this table.
    if !link_result.chunks.is_empty() {
        bootstrap.imports = link_result.chunks[0].imports.clone();
    }

    // Prepend bootstrap, shift all ref_func indices in component chunks by +1
    let mut all_chunks = vec![bootstrap];
    for chunk in &link_result.chunks {
        let mut adjusted = chunk.clone();
        let code = &mut adjusted.code;
        let mut ip = 0;
        while ip < code.len() {
            if let Some(op) = Op::from_byte(code[ip]) {
                match op {
                    Op::ref_func => {
                        if ip + 2 < code.len() {
                            let old_idx = ((code[ip + 1] as u16) << 8) | (code[ip + 2] as u16);
                            let new_idx = old_idx + 1;
                            code[ip + 1] = (new_idx >> 8) as u8;
                            code[ip + 2] = (new_idx & 0xff) as u8;
                        }
                        ip += 4; // op + u16 + upvalue_count
                        if ip - 1 < code.len() {
                            let uv_count = code[ip - 1] as usize;
                            ip += uv_count * 2;
                        }
                        continue;
                    }
                    Op::call_import => { ip += 4; continue; }
                    Op::r#const | Op::local_get | Op::local_set
                    | Op::global_get | Op::global_set
                    | Op::struct_get | Op::struct_set
                    | Op::array_new
                    | Op::br | Op::br_if_true | Op::br_if_false
                    | Op::r#loop => { ip += 3; continue; }
                    Op::call | Op::call_ref | Op::upvalue_get | Op::upvalue_set => { ip += 2; continue; }
                    _ => { ip += 1; }
                }
            } else {
                ip += 1;
            }
        }
        all_chunks.push(adjusted);
    }
    all_chunks
}

// ═══════════════════════════════════════════════════════════
// Two VB modules — both produce output
// ═══════════════════════════════════════════════════════════

#[test]
fn bootstrap_two_vb_modules() {
    let (mut vm, output) = setup_vm();

    let vb1 = vybe_parser_basic::parse_program("Console.WriteLine(\"Module 1\")").expect("VB1 parse");
    let chunks1 = vybe_compiler_vb::Compiler::new().compile(&vb1).expect("VB1 compile");
    let comp1 = vybe_compiler_common::components::build_component(
        "mod1", vybe_bytecode::component::Language::VB, chunks1);

    let vb2 = vybe_parser_basic::parse_program("Console.WriteLine(\"Module 2\")").expect("VB2 parse");
    let chunks2 = vybe_compiler_vb::Compiler::new().compile(&vb2).expect("VB2 compile");
    let comp2 = vybe_compiler_common::components::build_component(
        "mod2", vybe_bytecode::component::Language::VB, chunks2);

    let mut linker = vybe_bytecode::Linker::new();
    linker.register_host_from_vm(&vm);
    linker.add_component(comp1);
    linker.add_component(comp2);
    let link_result = linker.link().expect("Link failed");

    let all_chunks = build_bootstrap(&link_result, 2, Some(1));
    // Adjust resolved imports: build_bootstrap prepends a chunk, shifting all indices by +1
    let adjusted_imports: Vec<ImportTarget> = link_result.resolved_imports.iter().map(|t| {
        match t {
            ImportTarget::ChunkFn { chunk_index, arity } => ImportTarget::ChunkFn {
                chunk_index: chunk_index + 1,
                arity: *arity,
            },
            other => other.clone(),
        }
    }).collect();
    vm.run_linked(all_chunks, adjusted_imports).expect("Bootstrap run failed");

    assert_eq!(output.borrow().as_slice(), &["Module 1", "Module 2"]);
}

// ═══════════════════════════════════════════════════════════
// VB library + C# entry — library runs first
// ═══════════════════════════════════════════════════════════

#[test]
fn bootstrap_vb_library_cs_entry() {
    let (mut vm, output) = setup_vm();

    let vb_src = "Console.WriteLine(\"VB lib loaded\")";
    let vb_prog = vybe_parser_basic::parse_program(vb_src).expect("VB parse");
    let vb_chunks = vybe_compiler_vb::Compiler::new().compile(&vb_prog).expect("VB compile");
    let vb_comp = vybe_compiler_common::components::build_component(
        "vb_lib", vybe_bytecode::component::Language::VB, vb_chunks);

    let cs_src = "Console.WriteLine(\"C# entry\");";
    let cs_prog = vybe_parser_csharp::parse(cs_src).expect("C# parse");
    let cs_chunks = vybe_compiler_csharp::Compiler::new().compile(&cs_prog).expect("C# compile");
    let cs_comp = vybe_compiler_common::components::build_component(
        "cs_main", vybe_bytecode::component::Language::CSharp, cs_chunks);

    let mut linker = vybe_bytecode::Linker::new();
    linker.register_host_from_vm(&vm);
    linker.add_component(vb_comp);
    linker.add_component(cs_comp);
    let link_result = linker.link().expect("Link failed");

    // entry is comp index 1 (cs_main), library is comp 0 (vb_lib)
    let all_chunks = build_bootstrap(&link_result, 2, Some(1));
    // Adjust resolved imports: build_bootstrap prepends a chunk, shifting all indices by +1
    let adjusted_imports: Vec<ImportTarget> = link_result.resolved_imports.iter().map(|t| {
        match t {
            ImportTarget::ChunkFn { chunk_index, arity } => ImportTarget::ChunkFn {
                chunk_index: chunk_index + 1,
                arity: *arity,
            },
            other => other.clone(),
        }
    }).collect();
    vm.run_linked(all_chunks, adjusted_imports).expect("Bootstrap run failed");

    let out = output.borrow();
    assert_eq!(out[0], "VB lib loaded");
    assert_eq!(out[1], "C# entry");
}

// ═══════════════════════════════════════════════════════════
// Three languages: Ruby + PHP library, JS entry
// ═══════════════════════════════════════════════════════════

#[test]
fn bootstrap_three_languages() {
    let (mut vm, output) = setup_vm();

    let rb_src = "puts 'Ruby loaded'";
    let rb_prog = vybe_parser_ruby::parse(rb_src).expect("Ruby parse");
    let rb_chunks = vybe_compiler_ruby::Compiler::new().compile(&rb_prog).expect("Ruby compile");
    let rb_comp = vybe_compiler_common::components::build_component(
        "rb_lib", vybe_bytecode::component::Language::Ruby, rb_chunks);

    let php_src = "<?php echo 'PHP loaded';";
    let php_prog = vybe_parser_php::parse(php_src).expect("PHP parse");
    let php_chunks = vybe_compiler_php::Compiler::new().compile(&php_prog).expect("PHP compile");
    let php_comp = vybe_compiler_common::components::build_component(
        "php_lib", vybe_bytecode::component::Language::Php, php_chunks);

    let js_src = "console.log('JS entry');";
    vybe_compiler_js::register_js_coercion(&mut vm);
    let js_prog = vybe_parser_js::parse(js_src).expect("JS parse");
    let js_chunks = vybe_compiler_js::Compiler::new().compile(&js_prog).expect("JS compile");
    let js_comp = vybe_compiler_common::components::build_component(
        "js_main", vybe_bytecode::component::Language::JS, js_chunks);

    let mut linker = vybe_bytecode::Linker::new();
    linker.register_host_from_vm(&vm);
    linker.add_component(rb_comp);
    linker.add_component(php_comp);
    linker.add_component(js_comp);
    let link_result = linker.link().expect("Link failed");

    let all_chunks = build_bootstrap(&link_result, 3, Some(2));
    // Adjust resolved imports: build_bootstrap prepends a chunk, shifting all indices by +1
    let adjusted_imports: Vec<ImportTarget> = link_result.resolved_imports.iter().map(|t| {
        match t {
            ImportTarget::ChunkFn { chunk_index, arity } => ImportTarget::ChunkFn {
                chunk_index: chunk_index + 1,
                arity: *arity,
            },
            other => other.clone(),
        }
    }).collect();
    vm.run_linked(all_chunks, adjusted_imports).expect("Bootstrap run failed");

    let out = output.borrow();
    assert_eq!(out.len(), 3);
    assert_eq!(out[0], "Ruby loaded");
    assert_eq!(out[1], "PHP loaded");
    assert_eq!(out[2], "JS entry");
}

// ═══════════════════════════════════════════════════════════
// VB defines function, JS calls it — shared globals via bootstrap
// ═══════════════════════════════════════════════════════════

#[test]
fn bootstrap_vb_function_js_calls() {
    let (mut vm, output) = setup_vm();
    vybe_compiler_js::register_js_coercion(&mut vm);

    let vb_src = r#"
Public Function Square(x As Integer) As Integer
    Return x * x
End Function
"#;
    let vb_prog = vybe_parser_basic::parse_program(vb_src).expect("VB parse");
    let vb_chunks = vybe_compiler_vb::Compiler::new().compile(&vb_prog).expect("VB compile");
    let vb_comp = vybe_compiler_common::components::build_component(
        "math_lib", vybe_bytecode::component::Language::VB, vb_chunks);

    let js_src = r#"
var result = square(7);
console.log(result);
"#;
    let js_prog = vybe_parser_js::parse(js_src).expect("JS parse");
    let js_chunks = vybe_compiler_js::Compiler::new().compile(&js_prog).expect("JS compile");
    let js_comp = vybe_compiler_common::components::build_component(
        "main", vybe_bytecode::component::Language::JS, js_chunks);

    let mut linker = vybe_bytecode::Linker::new();
    linker.register_host_from_vm(&vm);
    linker.add_component(vb_comp);
    linker.add_component(js_comp);
    let link_result = linker.link().expect("Link failed");

    let all_chunks = build_bootstrap(&link_result, 2, Some(1));
    // Adjust resolved imports: build_bootstrap prepends a chunk, shifting all indices by +1
    let adjusted_imports: Vec<ImportTarget> = link_result.resolved_imports.iter().map(|t| {
        match t {
            ImportTarget::ChunkFn { chunk_index, arity } => ImportTarget::ChunkFn {
                chunk_index: chunk_index + 1,
                arity: *arity,
            },
            other => other.clone(),
        }
    }).collect();
    vm.run_linked(all_chunks, adjusted_imports).expect("Bootstrap run failed");

    assert_eq!(output.borrow().as_slice(), &["49"]);
}
