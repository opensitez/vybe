use vybe_parser_ruby::parse;
use vybe_compiler_ruby::Compiler;

fn compile_ok(src: &str) {
    let program = parse(src).expect("parse failed");
    let mut c = Compiler::new();
    let res = c.compile(&program);
    assert!(res.is_ok(), "compile failed for:\n{}\nerror: {:?}", src, res.err());
}

// ── DateTime (same as VB DateTime, PHP DateTime, Python datetime) ──
#[test] fn time_now() { compile_ok("t = Time.new"); }

// ── StringBuilder (same as VB/C# StringBuilder, PHP StringBuilder) ──
#[test] fn string_io() { compile_ok("sb = StringIO.new\nsb.write('hello')"); }

// ── Random (same as VB Random, Python random, PHP Random) ──
#[test] fn random_new() { compile_ok("rng = Random.new"); }

// ── Hash as dict (same as PHP arrays, Python dict, JS object) ──
#[test] fn hash_dict() { compile_ok("h = Hash.new"); }

// ── Set (same as Python set, C# HashSet) ──
#[test] fn set_new() { compile_ok("s = Set.new"); }

// ── Sockets (same as VB TcpClient, Python socket, PHP fsockopen) ──
#[test] fn tcp_socket() { compile_ok("sock = TCPSocket.new('localhost', 80)"); }
#[test] fn tcp_server() { compile_ok("server = TCPServer.new('0.0.0.0', 8080)"); }

// ── Process (same as VB Process, Python subprocess, PHP exec) ──
#[test] fn system_call() { compile_ok("system('ls')"); }

// ── Threading (same as Python threading, JS Worker) ──
#[test] fn thread_new() { compile_ok("t = Thread.new(1)"); }
#[test] fn mutex_new() { compile_ok("m = Mutex.new"); }

// ── Fiber (same as PHP Fiber, JS generator) ──
#[test] fn fiber_new() { compile_ok("f = Fiber.new(1)"); }

// ── Exception types (cross-language compatible) ──
#[test] fn exception_new() { compile_ok("e = RuntimeError.new('oops')"); }
#[test] fn type_error_new() { compile_ok("e = TypeError.new('bad type')"); }
#[test] fn argument_error() { compile_ok("e = ArgumentError.new('bad arg')"); }

// ── File operations (same as Python open, PHP fopen) ──
#[test] fn file_read() { compile_ok("content = 'test.txt'.read"); }

// ── Lambda / Proc (same as JS arrow functions, PHP closures) ──
#[test] fn lambda_compat() { compile_ok("add = -> (a, b) { a + b }\nresult = add.call(1, 2)"); }

// ── Enumerable (same as PHP array_map, Python map, JS Array.map) ──
#[test] fn map_compat() { compile_ok("doubled = [1, 2, 3].map { |x| x * 2 }"); }
#[test] fn filter_compat() { compile_ok("evens = [1, 2, 3, 4].select { |x| x % 2 == 0 }"); }
#[test] fn reduce_compat() { compile_ok("sum = [1, 2, 3].reduce(0) { |a, b| a + b }"); }

// ── Components (cross-language module) ──
#[test]
fn component_compile() {
    let src = r#"
def add(a, b)
  a + b
end

def multiply(a, b)
  a * b
end
"#;
    let program = parse(src).expect("parse failed");
    let result = vybe_compiler_ruby::compile_component(&program, "math_utils");
    assert!(result.is_ok(), "component compile failed: {:?}", result.err());
}
