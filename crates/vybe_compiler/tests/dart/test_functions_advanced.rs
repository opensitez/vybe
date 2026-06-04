use super::helpers::{compile_ok, run_prints};

// ── typedef ──────────────────────────────────────────────────

#[test]
fn typedef_simple() {
    compile_ok("typedef IntOp = int Function(int, int);");
}
#[test]
fn typedef_predicate() {
    compile_ok("typedef Predicate<T> = bool Function(T);");
}
#[test]
fn typedef_void() {
    compile_ok("typedef Callback = void Function(String);");
}
#[test]
fn typedef_no_params() {
    compile_ok("typedef Factory = Object Function();");
}

#[test]
fn typedef_used() {
    compile_ok(
        "typedef Transformer = String Function(String); String apply(String s, Transformer t) { return t(s); }",
    );
}

#[test]
fn typedef_as_field() {
    compile_ok("typedef Handler = void Function(String); class Server { Handler? onMessage; }");
}

// ── required named parameters ────────────────────────────────

#[test]
fn required_named() {
    compile_ok("void connect({required String host, required int port}) {}");
}
#[test]
fn required_named_called() {
    compile_ok("void f({required String name}) {} void main() { f(name: 'Alice'); }");
}

#[test]
fn required_mixed() {
    compile_ok("void f(int x, {required String name, int port = 80}) {}");
}

#[test]
fn required_in_class() {
    compile_ok(
        "class Server { final String host; final int port; Server({required this.host, required this.port}); }",
    );
}

#[test]
fn required_result() {
    let out = run_prints(
        r#"
void greet({required String name}) { print('Hello $name'); }
void main() { greet(name: 'Dart'); }
"#,
    );
    assert_eq!(out, ["Hello Dart"]);
}

// ── Functions as first-class values ─────────────────────────

#[test]
fn fn_as_variable() {
    compile_ok("int double(int x) { return x * 2; } var fn = double;");
}
#[test]
fn fn_as_arg() {
    compile_ok("void apply(int x, int Function(int) fn) { print(fn(x)); }");
}
#[test]
fn fn_return_fn() {
    compile_ok("Function makeAdder(int n) { return (int x) => x + n; }");
}
#[test]
fn fn_in_list() {
    compile_ok(
        "int double(int x) => x * 2; int triple(int x) => x * 3; var ops = [double, triple];",
    );
}

#[test]
fn fn_as_arg_result() {
    let out = run_prints(
        r#"
void apply(int x, int Function(int) fn) { print(fn(x)); }
void main() { apply(5, (x) => x * 2); }
"#,
    );
    assert_eq!(out, ["10"]);
}

#[test]
fn fn_return_fn_result() {
    let out = run_prints(
        r#"
Function makeAdder(int n) { return (int x) => x + n; }
void main() { var add5 = makeAdder(5); print(add5(3)); }
"#,
    );
    assert_eq!(out, ["8"]);
}

// ── Recursion ────────────────────────────────────────────────

#[test]
fn recursive_factorial() {
    compile_ok("int fact(int n) { return n <= 1 ? 1 : n * fact(n - 1); }");
}

#[test]
fn recursive_factorial_result() {
    let out = run_prints(
        r#"
int fact(int n) { return n <= 1 ? 1 : n * fact(n - 1); }
void main() { print(fact(5)); }
"#,
    );
    assert_eq!(out, ["120"]);
}

#[test]
fn recursive_fibonacci() {
    compile_ok("int fib(int n) { return n <= 1 ? n : fib(n - 1) + fib(n - 2); }");
}

#[test]
fn recursive_fibonacci_result() {
    let out = run_prints(
        r#"
int fib(int n) { return n <= 1 ? n : fib(n - 1) + fib(n - 2); }
void main() { print(fib(7)); }
"#,
    );
    assert_eq!(out, ["13"]);
}

#[test]
fn recursive_sum() {
    let out = run_prints(
        r#"
int sum(int n) { return n <= 0 ? 0 : n + sum(n - 1); }
void main() { print(sum(10)); }
"#,
    );
    assert_eq!(out, ["55"]);
}

// ── Closures ────────────────────────────────────────────────

#[test]
fn closure_capture() {
    compile_ok("void main() { var x = 10; var fn = () => x * 2; print(fn()); }");
}
#[test]
fn closure_mutate() {
    compile_ok(
        "void main() { var count = 0; var inc = () { count++; }; inc(); inc(); print(count); }",
    );
}

#[test]
fn closure_capture_result() {
    let out = run_prints("void main() { var x = 10; var fn = () => x * 2; print(fn()); }");
    assert_eq!(out, ["20"]);
}

#[test]
fn closure_counter() {
    let out = run_prints(
        r#"
void main() {
  var count = 0;
  var inc = () { count++; };
  inc(); inc(); inc();
  print(count);
}
"#,
    );
    assert_eq!(out, ["3"]);
}

#[test]
fn closure_in_loop() {
    compile_ok("void main() { var fns = []; for (var i = 0; i < 3; i++) { fns.add(() => i); } }");
}

// ── Arrow functions on methods ───────────────────────────────

#[test]
fn arrow_method() {
    compile_ok("class A { int double(int x) => x * 2; }");
}
#[test]
fn arrow_getter() {
    compile_ok("class A { int _x = 5; int get x => _x; }");
}
#[test]
fn arrow_top_level() {
    compile_ok("int square(int x) => x * x;");
}

#[test]
fn arrow_method_result() {
    let out = run_prints(
        r#"
class A { int double(int x) => x * 2; }
void main() { var a = A(); print(a.double(7)); }
"#,
    );
    assert_eq!(out, ["14"]);
}

#[test]
fn arrow_top_level_result() {
    let out = run_prints("int sq(int x) => x * x; void main() { print(sq(9)); }");
    assert_eq!(out, ["81"]);
}

// ── Optional positional params (more cases) ──────────────────

#[test]
fn optional_pos_unset() {
    let out = run_prints(
        r#"
void greet([String name = 'World']) { print('Hello $name'); }
void main() { greet(); }
"#,
    );
    assert_eq!(out, ["Hello World"]);
}

#[test]
fn optional_pos_set() {
    let out = run_prints(
        r#"
void greet([String name = 'World']) { print('Hello $name'); }
void main() { greet('Dart'); }
"#,
    );
    assert_eq!(out, ["Hello Dart"]);
}

#[test]
fn optional_multiple() {
    compile_ok(
        "String format(String s, [int width = 10, String fill = ' ']) { return s.padLeft(width, fill); }",
    );
}

// ── Named params (more cases) ────────────────────────────────

#[test]
fn named_unset() {
    let out = run_prints(
        r#"
void show({String msg = 'default'}) { print(msg); }
void main() { show(); }
"#,
    );
    assert_eq!(out, ["default"]);
}

#[test]
fn named_set() {
    let out = run_prints(
        r#"
void show({String msg = 'default'}) { print(msg); }
void main() { show(msg: 'custom'); }
"#,
    );
    assert_eq!(out, ["custom"]);
}

#[test]
fn named_order_independent() {
    compile_ok("void f({int a = 0, int b = 0}) {} void main() { f(b: 2, a: 1); }");
}

// ── Function type annotations ────────────────────────────────

#[test]
fn fn_type_annotation() {
    compile_ok("void Function(int) handler = (x) { print(x); };");
}
#[test]
fn fn_type_param() {
    compile_ok("T apply<T>(T val, T Function(T) fn) => fn(val);");
}
#[test]
fn fn_type_return() {
    compile_ok("int Function(int) doubler() { return (x) => x * 2; }");
}

// ── Multiple return values via records ───────────────────────

#[test]
fn return_record() {
    compile_ok("(int, int) minMax(List<int> list) { return (list.first, list.last); }");
}
#[test]
fn return_record_named() {
    compile_ok("({String name, int age}) person() { return (name: 'Alice', age: 30); }");
}

// ── void and null returns ────────────────────────────────────

#[test]
fn void_return() {
    compile_ok("void f() { return; }");
}
#[test]
fn null_return() {
    compile_ok("String? f() { return null; }");
}
