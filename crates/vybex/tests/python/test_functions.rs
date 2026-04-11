use super::helpers::{run_python, run_python_one, compile_ok};

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

// Generators (yield)

#[test] fn yield_basic() { compile_ok("def gen():\n    yield 1\n    yield 2\n    yield 3\n"); }
#[test] fn yield_no_value() { compile_ok("def gen():\n    yield\n"); }
#[test] fn yield_in_loop() { compile_ok("def count_up(n):\n    i = 0\n    while i < n:\n        yield i\n        i += 1\n"); }
#[test] fn yield_from() { compile_ok("def chain(a, b):\n    yield from a\n    yield from b\n"); }

// Runtime function tests

#[test]
fn function_call_runtime() {
    let out = run_python("def greet(name):\n    print(\"Hello, \" + name + \"!\")\ngreet(\"Python\")\n");
    assert_eq!(out, vec!["Hello, Python!"]);
}

#[test]
fn function_return_runtime() {
    assert_eq!(run_python_one("def add(a, b):\n    return a + b\nprint(add(3, 4))\n"), "7");
}

#[test]
fn nested_function_runtime() {
    assert_eq!(run_python_one("def outer():\n    def inner():\n        return 42\n    return inner()\nprint(outer())\n"), "42");
}

#[test]
fn recursive_function_runtime() {
    let out = run_python(r#"
def fib(n):
    if n <= 1:
        return n
    return fib(n - 1) + fib(n - 2)
print(fib(10))
"#);
    assert_eq!(out, vec!["55"]);
}

// map / filter

#[test]
fn map_basic() {
    compile_ok("result = list(map(lambda x: x * 2, [1, 2, 3]))\nprint(result)\n");
}

#[test]
fn map_with_named_func() {
    compile_ok("def double(x):\n    return x * 2\nresult = map(double, [1, 2, 3])\n");
}

#[test]
fn filter_basic() {
    compile_ok("result = list(filter(lambda x: x > 0, [-1, 0, 1, 2]))\nprint(result)\n");
}

#[test]
fn filter_with_named_func() {
    compile_ok("def is_even(x):\n    return x % 2 == 0\nresult = filter(is_even, [1, 2, 3, 4])\n");
}
