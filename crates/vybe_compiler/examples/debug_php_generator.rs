use std::sync::{Arc, Mutex};

use vybe_bytecode::{HostContext, VM, Value};
use vybe_compiler::compiler::Compiler;
use vybe_compiler::languages::php;
use vybe_compiler::profile::parse_profile;

fn main() {
    let src = r#"<?php
function accumulator() {
    $total = 0;
    while (true) {
        $value = yield $total;
        if ($value === null) break;
        $total += $value;
    }
}
$gen = accumulator();
echo "current=".$gen->current().";";
echo "send10=".$gen->send(10).";";
echo "send20=".$gen->send(20).";";
echo "send30=".$gen->send(30).";";
"#;

    let module = php::parse(src).expect("parse");
    let profile = parse_profile(php::profile_source()).expect("profile");
    let chunks = Compiler::with_profile(profile)
        .compile(&module)
        .expect("compile");

    for (index, chunk) in chunks.iter().enumerate() {
        if !chunk.name.contains("accumulator") && chunk.name != "<script>" {
            continue;
        }
        println!(
            "\n-- chunk {index}: {} locals={} --",
            chunk.name, chunk.local_count
        );
        for constant_index in [48usize, 55, 62, 70, 73, 74, 75, 78, 79, 80] {
            if let Some(value) = chunk.constants.get(constant_index) {
                println!("const[{constant_index}]={value}");
            }
        }
        let disasm = vybe_bytecode::debug::disassemble(chunk);
        let lines: Vec<_> = disasm.lines().collect();
        for (line_index, line) in lines.iter().enumerate() {
            if line.contains("suspend") || line.contains("resume") {
                let start = line_index.saturating_sub(40);
                let end = (line_index + 320).min(lines.len());
                for selected in &lines[start..end] {
                    println!("{selected}");
                }
            }
        }
    }

    let mut vm = VM::new();
    let output: Arc<Mutex<Vec<String>>> = Arc::new(Mutex::new(Vec::new()));
    let out = output.clone();
    vybe_host::register_all(&mut vm);
    vm.register_host_fn(
        "wasi:cli",
        "log",
        Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
            let joined = args
                .iter()
                .map(|arg| format!("{}", arg))
                .collect::<Vec<_>>()
                .join(" ");
            out.lock().unwrap().push(joined);
            Value::Null
        }),
    );
    vybe_host::setup_namespaces(&mut vm);
    let result = vm.run(chunks);
    println!("run={result:?}");
    println!("output={:?}", output.lock().unwrap());
}
