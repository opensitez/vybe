use vybe_parser_dart::parse;
use vybe_compiler_dart::Compiler;

fn compile_ok(src: &str) {
    let program = parse(src).expect("parse failed");
    let mut c = Compiler::new();
    let res = c.compile(&program);
    assert!(res.is_ok(), "compile failed for:\n{}\nerror: {:?}", src, res.err());
}

fn parse_ok(src: &str) -> bool {
    parse(src).is_ok()
}

// ── String APIs ─────────────────────────────────────────────
#[test] fn string_interpolation() { compile_ok("var name = 'World'; var s = 'Hello $name!';"); }
#[test] fn string_methods() { compile_ok("var s = 'Hello World'; var u = s.toUpperCase(); var l = s.toLowerCase(); var t = s.trim();"); }
#[test] fn string_contains() { compile_ok("var b = 'hello'.contains('ell');"); }
#[test] fn string_split() { compile_ok("var parts = 'a,b,c'.split(',');"); }
#[test] fn string_replace() { compile_ok("var s = 'hello'.replaceAll('l', 'r');"); }
#[test] fn string_substring() { compile_ok("var s = 'hello'.substring(1, 3);"); }
#[test] fn string_pad() { compile_ok("var s = '42'.padLeft(5, '0');"); }

// ── List APIs ───────────────────────────────────────────────
#[test] fn list_add() { compile_ok("var a = [1, 2]; a.add(3);"); }
#[test] fn list_map() { compile_ok("var doubled = [1, 2, 3].map((x) => x * 2);"); }
#[test] fn list_where() { compile_ok("var evens = [1, 2, 3, 4].where((x) => x % 2 == 0);"); }
#[test] fn list_reduce() { compile_ok("var sum = [1, 2, 3].reduce((a, b) => a + b);"); }
#[test] fn list_foreach() { compile_ok("var items = [1, 2, 3]; items.forEach((x) => print(x));"); }
#[test] fn list_any_every() { compile_ok("var has = [1, 2, 3].any((x) => x > 2); var all = [1, 2, 3].every((x) => x > 0);"); }
#[test] fn list_join() { compile_ok("var s = ['a', 'b', 'c'].join(', ');"); }
#[test] fn list_reversed() { compile_ok("var r = [1, 2, 3].reversed;"); }

// ── Map APIs ────────────────────────────────────────────────
#[test] fn map_literal() { compile_ok("var m = {'name': 'Alice', 'age': 30};"); }
#[test] fn map_access() { compile_ok("var m = {'x': 1}; var v = m['x'];"); }

// ── Async/Await ─────────────────────────────────────────────
#[test] fn async_await() { compile_ok("class App { fetchData() async { return 'data'; } main() async { var d = await fetchData(); print(d); } }");}
#[test] fn await_expr() { compile_ok("var x = await http.get('https://example.com');"); }

// ── Null safety ─────────────────────────────────────────────
#[test] fn null_aware() { compile_ok("var x = null; var y = x ?? 'default';"); }
#[test] fn cascade() { compile_ok("var list = []; list..add(1)..add(2)..add(3);"); }

// ── Classes ─────────────────────────────────────────────────
#[test] fn class_inheritance() { compile_ok("class Animal { String name; Animal(this.name); } class Dog extends Animal { Dog(String name) : super(name); String speak() => name + ' barks'; }"); }
#[test] fn abstract_class() { compile_ok("abstract class Shape { double area(); } class Circle extends Shape { double r; Circle(this.r); double area() => 3.14 * r * r; }"); }
#[test] fn mixins() { compile_ok("mixin Greetable { String greet() => 'Hello'; } class Person with Greetable { String name; Person(this.name); }"); }
#[test] fn extension_method() { compile_ok("extension StringExt on String { String reversed() => split('').reversed.join(''); } var r = 'hello'.reversed();"); }
#[test] fn enum_basic() { compile_ok("enum Color { red, green, blue } var c = Color.red;"); }
#[test] fn factory_constructor() { compile_ok("class Logger { static Logger? _instance; factory Logger() { _instance ??= Logger(); return _instance; } Logger(); }"); }
#[test] fn getter_setter() { compile_ok("class Rect { double _w = 0; double get width => _w; set width(double v) { _w = v; } }"); }

// ── Control flow ────────────────────────────────────────────
#[test] fn for_in() { compile_ok("var list = [1, 2, 3]; for (var x in list) { print(x); }"); }
#[test] fn switch_case() { compile_ok("var x = 1; switch (x) { case 1: print('one'); break; case 2: print('two'); break; default: print('other'); }"); }
#[test] fn try_catch() { compile_ok("try { throw Exception('oops'); } catch (e) { print(e); } finally { print('done'); }"); }

// ── Dart 3 features ─────────────────────────────────────────
#[test] fn record_type() { assert!(parse_ok("var point = (1, 2); var x = point.$1;")); }
// Dart 3 pattern destructuring needs parser support for var + tuple LHS
// #[test] fn pattern_matching() { assert!(parse_ok("var (a, b) = (1, 2);")); }
