use vybec::parser_php::Parser;
use vybec::compiler_php::Compiler;
use vybe_bytecode::{VM, Value};
use std::sync::{Arc, Mutex};

pub fn parse(src: &str) -> vybec::parser_php::Program {
    Parser::new(src).expect("lexer failed").parse_program().expect("parse failed")
}

pub fn compile_ok(src: &str) {
    let program = parse(src);
    let res = Compiler::new().compile(&program);
    assert!(res.is_ok(), "compile failed: {:?}", res.err());
    assert!(!res.unwrap().is_empty());
}

pub fn compile(src: &str) -> Vec<vybe_bytecode::Chunk> {
    let program = parse(src);
    Compiler::new().compile(&program).expect("compile failed")
}

pub fn run(src: &str) -> Value {
    let chunks = compile(src);
    let mut vm = VM::new();
    vybe_host::register_all(&mut vm);
    vm.run(chunks).unwrap()
}

pub fn run_prints(src: &str) -> Vec<String> {
    let chunks = compile(src);
    let mut vm = VM::new();
    vybe_host::register_all(&mut vm);
    let output = Arc::new(Mutex::new(Vec::<String>::new()));
    let out = output.clone();
    vm.register_host_fn("wasi:cli", "log", Box::new(move |_ctx: &mut vybe_bytecode::HostContext, args: &[Value]| {
        let s: Vec<String> = args.iter().map(|a| format!("{}", a)).collect();
        out.lock().unwrap().push(s.join(" "));
        Value::Null
    }));
    vm.run(chunks).unwrap();
    output.lock().unwrap().clone()
}
