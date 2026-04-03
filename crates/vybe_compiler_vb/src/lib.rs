pub mod compiler;
pub mod scope;
mod statements;
mod expressions;
mod builtins;
mod classes;

pub use compiler::Compiler;
use vybe_bytecode::{VM, Value};
use std::rc::Rc;
use std::cell::RefCell;

/// Set up VM with all host functions needed by VB, then compile.
pub fn setup_and_compile(
    vm: &mut VM,
    program: &vybe_parser_basic::ast::Program,
) -> Result<Vec<vybe_bytecode::Chunk>, String> {
    vybe_host::register_all(vm);
    Compiler::new().compile(program)
}

/// Set up VM with all host functions + GUI, then compile.
pub fn setup_and_compile_with_gui(
    vm: &mut VM,
    program: &vybe_parser_basic::ast::Program,
    queue: Rc<RefCell<vybe_host::SideEffectQueue>>,
) -> Result<Vec<vybe_bytecode::Chunk>, String> {
    vybe_host::register_all_with_gui(vm, queue);
    Compiler::new().compile(program)
}

/// Convenience: parse VB source + compile + run, return captured output.
pub fn compile_and_run(source: &str) -> Result<Vec<String>, String> {
    let program = vybe_parser_basic::parse_program(source)
        .map_err(|e| format!("Parse error: {e}"))?;
    let mut vm = VM::new();
    let output: Rc<RefCell<Vec<String>>> = Rc::new(RefCell::new(Vec::new()));
    let out = output.clone();

    // Register host functions
    vybe_host::register_all(&mut vm);

    // Override console.log to capture output
    vm.register_host_fn("wasi:cli", "log", Box::new(move |_vm: &mut VM, args: &[Value]| {
        let parts: Vec<String> = args.iter().map(|v| format!("{v}")).collect();
        out.borrow_mut().push(parts.join(" "));
        Value::Null
    }));

    let chunks = Compiler::new().compile(&program)?;
    vm.run(chunks).map_err(|e| format!("Runtime error: {e}"))?;
    Ok(output.borrow().clone())
}
