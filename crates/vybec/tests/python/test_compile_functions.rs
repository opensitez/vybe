use vybec::parser_python::parse;
use vybec::compiler_python::Compiler;

fn compile_ok(src: &str) {
    let module = parse(src).expect("parse failed");
    let mut c = Compiler::new();
    let res = c.compile(&module);
    assert!(res.is_ok(), "compile failed: {:?}", res.err());
}

#[test]
fn simple_function() {
    compile_ok("def greet():\n    print('hello')\ngreet()\n");
}

#[test]
fn function_with_args() {
    compile_ok("def add(a, b):\n    return a + b\nresult = add(3, 4)\n");
}

#[test]
fn function_with_defaults() {
    compile_ok("def greet(name, greeting='hello'):\n    print(greeting, name)\ngreet('world')\n");
}

#[test]
fn nested_function() {
    compile_ok("def outer():\n    def inner():\n        return 42\n    return inner()\n");
}

#[test]
fn lambda_simple() {
    compile_ok("f = lambda x: x + 1\nprint(f(5))\n");
}

#[test]
fn lambda_multi_arg() {
    compile_ok("f = lambda x, y: x + y\nprint(f(1, 2))\n");
}

#[test]
fn class_basic() {
    compile_ok("class Foo:\n    def bar(self):\n        print('bar')\n");
}

#[test]
fn class_with_init() {
    compile_ok("class Dog:\n    def __init__(self, name):\n        self.name = name\n    def bark(self):\n        print(self.name)\n");
}

#[test]
fn list_comprehension() {
    compile_ok("squares = [x * x for x in range(10)]\nprint(squares)\n");
}

#[test]
fn list_comp_with_filter() {
    compile_ok("evens = [x for x in range(20) if x % 2 == 0]\n");
}

#[test]
fn dict_comp() {
    compile_ok("d = {k: k * 2 for k in range(5)}\n");
}

#[test]
fn method_calls() {
    compile_ok("x = [1, 2]\nx.append(3)\nprint(len(x))\n");
}

#[test]
fn string_methods() {
    compile_ok("s = \"Hello World\"\nprint(s.upper())\nprint(s.lower())\nprint(s.split())\n");
}

#[test]
fn builtins() {
    compile_ok("print(len([1,2,3]))\nprint(str(42))\nprint(int('5'))\nprint(float('3.14'))\nprint(abs(-5))\n");
}
