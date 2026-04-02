pub mod compiler;
pub mod scope;

pub use compiler::Compiler;

use vybe_bytecode::VM;

/// Set up VM with all host functions needed by PHP, then compile.
pub fn setup_and_compile(
    vm: &mut VM,
    program: &vybe_parser_php::Program,
) -> Result<Vec<vybe_bytecode::Chunk>, String> {
    vybe_host::register_all(vm);
    Compiler::new().compile(program)
}
