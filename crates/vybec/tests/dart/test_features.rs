use vybec::parser_dart::parse;
use vybec::compiler_dart::Compiler;

fn compile_ok(src: &str) {
    let program = parse(src).expect("parse failed");
    let mut c = Compiler::new();
    let res = c.compile(&program);
    assert!(res.is_ok(), "compile failed for:\n{}\nerror: {:?}", src, res.err());
}

// ── Classes ────────────────────────────────────────────────

#[test] fn class_basic() {
    compile_ok("class Dog { String name; Dog(this.name); }");
}
#[test] fn class_method() {
    compile_ok("class Dog { String name; Dog(this.name); String bark() { return 'Woof!'; } }");
}
#[test] fn class_static_method() {
    compile_ok("class Math { static int add(int a, int b) { return a + b; } }");
}
#[test] fn class_getter() {
    compile_ok("class Rect { int w; int h; Rect(this.w, this.h); get area { return w * h; } }");
}
#[test] fn class_getter_typed() {
    compile_ok("class Rect { int w; int h; Rect(this.w, this.h); int get area { return w * h; } }");
}
#[test] fn class_setter() {
    compile_ok("class Box { int _size = 0; set size(int v) { _size = v; } }");
}
#[test] fn class_inheritance() {
    compile_ok("class Animal { String name; Animal(this.name); } class Dog extends Animal { Dog(String n) : super(n); }");
}
#[test] fn class_named_constructor() {
    compile_ok("class Point { int x; int y; Point(this.x, this.y); Point.origin() : this(0, 0); }");
}

// ── Collection methods ─────────────────────────────────────

#[test] fn list_map() {
    compile_ok("var x = [1, 2, 3].map((e) => e * 2);");
}
#[test] fn list_where() {
    compile_ok("var x = [1, 2, 3, 4].where((e) => e > 2);");
}
#[test] fn list_foreach() {
    compile_ok("void main() { [1, 2, 3].forEach((e) { print(e); }); }");
}
#[test] fn list_reduce() {
    compile_ok("var sum = [1, 2, 3, 4].reduce((a, b) => a + b);");
}
#[test] fn list_any() {
    compile_ok("var has = [1, 2, 3].any((e) => e > 2);");
}
#[test] fn list_every() {
    compile_ok("var all = [1, 2, 3].every((e) => e > 0);");
}
#[test] fn list_to_list() {
    compile_ok("var x = [1, 2, 3].toList();");
}
#[test] fn list_chain() {
    compile_ok("var x = [1, 2, 3].map((e) => e * 2).where((e) => e > 2).toList();");
}

// ── List properties ────────────────────────────────────────

#[test] fn list_length() {
    compile_ok("var n = [1, 2, 3].length;");
}
#[test] fn list_is_empty() {
    compile_ok("var b = [].isEmpty;");
}
#[test] fn list_is_not_empty() {
    compile_ok("var b = [1].isNotEmpty;");
}
#[test] fn list_first() {
    compile_ok("var f = [1, 2, 3].first;");
}
#[test] fn list_last() {
    compile_ok("var l = [1, 2, 3].last;");
}

// ── String methods ─────────────────────────────────────────

#[test] fn string_upper() { compile_ok("var x = 'hello'.toUpperCase();"); }
#[test] fn string_lower() { compile_ok("var x = 'HELLO'.toLowerCase();"); }
#[test] fn string_trim() { compile_ok("var x = '  hi  '.trim();"); }
#[test] fn string_split() { compile_ok("var x = 'a,b,c'.split(',');"); }
#[test] fn string_contains() { compile_ok("var x = 'hello'.contains('ell');"); }
#[test] fn string_starts_with() { compile_ok("var x = 'hello'.startsWith('he');"); }
#[test] fn string_replace() { compile_ok("var x = 'hello'.replaceAll('l', 'r');"); }

// ── Array methods ──────────────────────────────────────────

#[test] fn array_add() { compile_ok("var x = [1, 2]; x.add(3);"); }
#[test] fn array_join() { compile_ok("var x = [1, 2, 3].join(',');"); }
#[test] fn array_reversed() { compile_ok("var x = [1, 2, 3].reversed;"); }

// ── Control flow ───────────────────────────────────────────

#[test] fn if_else() { compile_ok("void main() { if (true) { print(1); } else { print(2); } }"); }
#[test] fn for_loop() { compile_ok("void main() { for (var i = 0; i < 10; i++) { print(i); } }"); }
#[test] fn for_in() { compile_ok("void main() { for (var x in [1, 2, 3]) { print(x); } }"); }
#[test] fn while_loop() { compile_ok("void main() { var i = 0; while (i < 5) { i = i + 1; } }"); }
#[test] fn switch_stmt() { compile_ok("void main() { var x = 1; switch(x) { case 1: print('one'); break; default: print('other'); } }"); }
#[test] fn try_catch() { compile_ok("void main() { try { throw 'error'; } catch (e) { print(e); } }"); }

// ── Functions ──────────────────────────────────────────────

#[test] fn lambda() { compile_ok("var f = (x) => x * 2;"); }
#[test] fn named_params() { compile_ok("void greet({String name = 'World'}) { print(name); }"); }
#[test] fn optional_params() { compile_ok("void greet([String name = 'World']) { print(name); }"); }

// ── Null safety ────────────────────────────────────────────

#[test] fn null_coalesce() { compile_ok("var x = null ?? 42;"); }
#[test] fn null_aware_access() { compile_ok("class A { int x = 1; } void main() { var a = null; var b = a?.x; }"); }

// ── Enums ──────────────────────────────────────────────────

#[test] fn enum_basic() { compile_ok("enum Color { red, green, blue }"); }

// ── String interpolation ───────────────────────────────────

#[test] fn string_interpolation() { compile_ok("var name = 'World'; var s = 'Hello $name!';"); }
#[test] fn string_interpolation_expr() { compile_ok("var x = 42; var s = 'Value is ${x + 1}';"); }

// ── Inheritance & super ────────────────────────────────────

#[test] fn super_constructor() {
    compile_ok("class Animal { String name; Animal(this.name); } class Dog extends Animal { Dog(String n) : super(n); }");
}
#[test] fn super_method_call() {
    compile_ok("class A { String greet() { return 'hello'; } } class B extends A { String greet() { return super.greet() + ' world'; } }");
}
#[test] fn inherited_method() {
    compile_ok("class A { int value() { return 42; } } class B extends A {} void main() { var b = B(); print(b.value()); }");
}
#[test] fn static_inheritance() {
    compile_ok("class Base { static int x() { return 1; } } class Child extends Base {}");
}
#[test] fn mixin_method() {
    compile_ok("class Printable { String describe() { return 'printable'; } } class Foo with Printable {}");
}

// ── Static methods ─────────────────────────────────────────

#[test] fn static_method_call() {
    compile_ok("class Utils { static int double(int x) { return x * 2; } } void main() { var r = Utils.double(5); }");
}

// ── Getters & setters ──────────────────────────────────────

#[test] fn getter_and_setter() {
    compile_ok("class Box { int _v = 0; get value { return _v; } set value(int v) { _v = v; } }");
}
#[test] fn getter_arrow_body() {
    compile_ok("class Circle { double r; Circle(this.r); get area { return 3.14 * r * r; } }");
}

// ── Higher-order collection chaining ───────────────────────

#[test] fn map_where_chain() {
    compile_ok("void main() { var x = [1,2,3,4,5].map((e) => e * 2).where((e) => e > 4).toList(); }");
}
#[test] fn reduce_sum() {
    compile_ok("void main() { var total = [1,2,3,4,5].reduce((a, b) => a + b); }");
}
#[test] fn any_every() {
    compile_ok("void main() { var h = [1,2,3].any((e) => e > 2); var a = [1,2,3].every((e) => e > 0); }");
}
#[test] fn foreach_side_effect() {
    compile_ok("void main() { [1,2,3].forEach((e) { print(e); }); }");
}

// ── Operator overloading ───────────────────────────────────

#[test] fn operator_plus_no_return_type() {
    compile_ok("class Vec { int x; int y; Vec(this.x, this.y); operator +(Vec other) { return Vec(x + other.x, y + other.y); } }");
}
#[test] fn operator_plus() {
    compile_ok("class Vec { int x; int y; Vec(this.x, this.y); Vec operator +(Vec other) { return Vec(x + other.x, y + other.y); } }");
}
#[test] fn operator_equals() {
    compile_ok("class Point { int x; int y; Point(this.x, this.y); bool operator ==(Object other) { return true; } }");
}
#[test] fn operator_index() {
    compile_ok("class Grid { List data; Grid(this.data); int operator [](int i) { return data[i]; } }");
}
#[test] fn operator_less_than() {
    compile_ok("class Score { int v; Score(this.v); bool operator <(Score other) { return v < other.v; } }");
}
#[test] fn operator_multiple() {
    compile_ok("class Num { int v; Num(this.v); Num operator +(Num o) { return Num(v + o.v); } Num operator -(Num o) { return Num(v - o.v); } Num operator *(Num o) { return Num(v * o.v); } }");
}
