use super::helpers::{compile_ok, run_prints};

// ── Generic classes ─────────────────────────────────────────

#[test] fn generic_box() { compile_ok("class Box<T> { T value; Box(this.value); }"); }
#[test] fn generic_pair() { compile_ok("class Pair<A, B> { A first; B second; Pair(this.first, this.second); }"); }
#[test] fn generic_stack() {
    compile_ok(r#"
class Stack<T> {
  List<T> _items = [];
  void push(T item) { _items.add(item); }
  T pop() { return _items.removeLast(); }
  bool get isEmpty => _items.isEmpty;
}
"#);
}

#[test] fn generic_optional() {
    compile_ok("class Maybe<T> { T? value; Maybe([this.value]); bool get hasValue => value != null; }");
}

#[test] fn generic_two_type_params() {
    compile_ok("class Result<T, E> { T? success; E? error; Result.ok(this.success); Result.err(this.error); }");
}

#[test] fn generic_box_used() {
    compile_ok("class Box<T> { T value; Box(this.value); } void main() { var b = Box(42); print(b.value); }");
}

#[test] fn generic_box_string() {
    compile_ok("class Box<T> { T value; Box(this.value); } void main() { var b = Box('hello'); }");
}

#[test] fn generic_pair_result() {
    let out = run_prints(r#"
class Pair<A, B> { A first; B second; Pair(this.first, this.second); }
void main() { var p = Pair(1, 'one'); print(p.first); }
"#);
    assert_eq!(out, ["1"]);
}

// ── Generic methods ──────────────────────────────────────────

#[test] fn generic_fn() { compile_ok("T identity<T>(T val) => val;"); }
#[test] fn generic_fn_two() { compile_ok("B transform<A, B>(A val, B Function(A) fn) => fn(val);"); }
#[test] fn generic_fn_list() { compile_ok("List<T> repeat<T>(T val, int times) => List.generate(times, (_) => val);"); }

#[test] fn generic_fn_result() {
    let out = run_prints("T id<T>(T v) => v; void main() { print(id(42)); }");
    assert_eq!(out, ["42"]);
}

#[test] fn generic_fn_string_result() {
    let out = run_prints("T id<T>(T v) => v; void main() { print(id('hello')); }");
    assert_eq!(out, ["hello"]);
}

// ── Typed collections ─────────────────────────────────────────

#[test] fn list_int_typed() { compile_ok("List<int> nums = [1, 2, 3];"); }
#[test] fn list_string_typed() { compile_ok("List<String> names = ['Alice', 'Bob'];"); }
#[test] fn list_bool_typed() { compile_ok("List<bool> flags = [true, false, true];"); }
#[test] fn map_string_int() { compile_ok("Map<String, int> scores = {'Alice': 90, 'Bob': 85};"); }
#[test] fn map_int_list() { compile_ok("Map<int, List<String>> grouped = {1: ['a', 'b'], 2: ['c']};"); }
#[test] fn set_int_typed() { compile_ok("Set<int> nums = {1, 2, 3};"); }
#[test] fn set_string_typed() { compile_ok("Set<String> tags = {'dart', 'flutter'};"); }

#[test] fn typed_list_operations() {
    compile_ok("List<int> nums = [3, 1, 2]; nums.sort(); var first = nums.first;");
}

// ── Generic constraints ──────────────────────────────────────

#[test] fn generic_extends() {
    compile_ok("class Container<T extends Object> { T value; Container(this.value); }");
}

#[test] fn generic_num_constraint() {
    compile_ok("T add<T extends num>(T a, T b) => (a + b) as T;");
}

#[test] fn generic_comparable() {
    compile_ok("T max<T extends Comparable<T>>(T a, T b) => a.compareTo(b) > 0 ? a : b;");
}

// ── Generic inheritance ──────────────────────────────────────

#[test] fn generic_extends_class() {
    compile_ok("class Base<T> { T value; Base(this.value); } class Child<T> extends Base<T> { Child(T v) : super(v); }");
}

#[test] fn generic_concrete_child() {
    compile_ok("class Base<T> { T value; Base(this.value); } class IntBox extends Base<int> { IntBox(int v) : super(v); }");
}

// ── Generic with null safety ─────────────────────────────────

#[test] fn generic_nullable_field() {
    compile_ok("class Cache<T> { T? _value; T? get value => _value; }");
}

#[test] fn generic_nullable_param() {
    compile_ok("T? tryParse<T>(String s, T? Function(String) parser) => parser(s);");
}

// ── Generic class methods ────────────────────────────────────

#[test] fn generic_method_on_class() {
    compile_ok(r#"
class Converter<T> {
  List<T> items;
  Converter(this.items);
  List<R> convert<R>(R Function(T) fn) => items.map(fn).toList();
}
"#);
}

// ── Generic in typedef ───────────────────────────────────────

#[test] fn generic_typedef() { compile_ok("typedef Mapper<T, R> = R Function(T);"); }
#[test] fn generic_typedef_used() {
    compile_ok("typedef Predicate<T> = bool Function(T); bool check<T>(T val, Predicate<T> pred) => pred(val);");
}

// ── Built-in generic usage ───────────────────────────────────

#[test] fn future_typed() { compile_ok("Future<int> f = Future.value(42);"); }
#[test] fn iterable_typed() { compile_ok("Iterable<int> it = [1, 2, 3];"); }
