use super::helpers::assert_js;

macro_rules! case {
    ($src:expr, [$($expected:expr),* $(,)?]) => {
        assert_js($src, &[$($expected),*]);
    };
}

#[test]
fn call_uses_explicit_receiver_for_plain_function() {
    case!(r#"
function label(prefix) {
    return prefix + ":" + this.name;
}
const obj = { name: "Ada" };
console.log(label.call(obj, "hi"));
"#, ["hi:Ada"]);
}

#[test]
fn call_can_override_object_literal_method_receiver() {
    case!(r#"
const left = {
    name: "left",
    show() {
        return this.name;
    }
};
const right = { name: "right" };
console.log(left.show.call(right));
"#, ["right"]);
}

#[test]
fn call_can_forward_multiple_arguments() {
    case!(r#"
function sum(a, b, c) {
    return this.base + a + b + c;
}
const ctx = { base: 1 };
console.log(sum.call(ctx, 2, 3, 4));
"#, ["10"]);
}

#[test]
fn call_can_borrow_hasownproperty() {
    case!(r#"
const hasOwn = Object.prototype.hasOwnProperty;
const obj = { x: 1 };
console.log(hasOwn.call(obj, "x"));
console.log(hasOwn.call(obj, "toString"));
"#, ["true", "false"]);
}

#[test]
fn call_can_borrow_array_join_for_array_like() {
    case!(r#"
const arrLike = { 0: "a", 1: "b", length: 2 };
console.log(Array.prototype.join.call(arrLike, "|"));
"#, ["a|b"]);
}

#[test]
fn call_receiver_is_visible_in_default_parameter_expression() {
    case!(r#"
function describe(prefix = this.tag) {
    console.log(prefix + ":" + this.tag);
}
describe.call({ tag: "ctx" });
"#, ["ctx:ctx"]);
}

#[test]
fn call_can_return_object_built_from_receiver_state() {
    case!(r#"
function make(key) {
    return { seen: this[key] };
}
const ctx = { value: 7 };
console.log(make.call(ctx, "value").seen);
"#, ["7"]);
}

#[test]
fn call_can_borrow_string_slice_for_primitive_this() {
    case!(r#"
console.log(String.prototype.slice.call("hello", 1, 4));
"#, ["ell"]);
}

#[test]
fn apply_spreads_array_arguments_into_call() {
    case!(r#"
function add(a, b, c) {
    return a + b + c;
}
console.log(add.apply(null, [1, 2, 3]));
"#, ["6"]);
}

#[test]
fn apply_uses_receiver_for_method_like_function() {
    case!(r#"
function tag(a, b) {
    return this.prefix + a + b;
}
const ctx = { prefix: "p:" };
console.log(tag.apply(ctx, ["x", "y"]));
"#, ["p:xy"]);
}

#[test]
fn apply_accepts_empty_argument_list() {
    case!(r#"
function size() {
    console.log(arguments.length);
}
size.apply(null, []);
"#, ["0"]);
}

#[test]
fn apply_can_forward_arguments_object() {
    case!(r#"
function target(a, b, c) {
    console.log(a + b + c);
}
function wrapper() {
    target.apply(null, arguments);
}
wrapper(2, 3, 4);
"#, ["9"]);
}

#[test]
fn apply_can_borrow_array_push_for_array_like() {
    case!(r#"
const obj = { length: 0 };
Array.prototype.push.apply(obj, ["a", "b"]);
console.log(obj.length);
console.log(obj[0] + obj[1]);
"#, ["2", "ab"]);
}

#[test]
fn apply_can_borrow_math_max_from_array() {
    case!(r#"
console.log(Math.max.apply(null, [3, 10, 5]));
"#, ["10"]);
}

#[test]
fn apply_receiver_used_in_default_parameter() {
    case!(r#"
function greet(name = this.name) {
    return "hi " + name;
}
console.log(greet.apply({ name: "Ada" }, []));
"#, ["hi Ada"]);
}

#[test]
fn apply_with_sparse_array_passes_undefined_holes() {
    case!(r#"
function pick(a, b) {
    console.log(a === undefined);
    console.log(b);
}
const args = [];
args[1] = "x";
pick.apply(null, args);
"#, ["true", "x"]);
}

#[test]
fn bind_prepends_single_argument() {
    case!(r#"
function add(a, b) {
    return a + b;
}
const inc = add.bind(null, 1);
console.log(inc(4));
"#, ["5"]);
}

#[test]
fn bind_prepends_multiple_arguments() {
    case!(r#"
function list(a, b, c) {
    return [a, b, c].join(",");
}
const head = list.bind(null, "a", "b");
console.log(head("c"));
"#, ["a,b,c"]);
}

#[test]
fn bind_preserves_first_bound_receiver() {
    case!(r#"
function show() {
    return this.name;
}
const bound = show.bind({ name: "first" });
console.log(bound.call({ name: "second" }));
"#, ["first"]);
}

#[test]
fn bind_on_extracted_method_restores_receiver() {
    case!(r#"
const obj = {
    name: "item",
    show(suffix) {
        return this.name + suffix;
    }
};
const fn = obj.show.bind(obj);
console.log(fn("!"));
"#, ["item!"]);
}

#[test]
fn bind_can_chain_bound_arguments() {
    case!(r#"
function parts(a, b, c) {
    return [a, b, c].join(":");
}
const one = parts.bind(null, "x");
const two = one.bind(null, "y");
console.log(two("z"));
"#, ["x:y:z"]);
}

#[test]
fn bind_chained_receiver_stays_from_first_bind() {
    case!(r#"
function show() {
    return this.name;
}
const one = show.bind({ name: "a" });
const two = one.bind({ name: "b" });
console.log(two());
"#, ["a"]);
}

#[test]
fn bind_length_reduces_by_bound_arguments() {
    case!(r#"
function sample(a, b, c) {}
const bound = sample.bind(null, 1);
console.log(bound.length);
"#, ["2"]);
}

#[test]
fn bind_length_clamps_at_zero() {
    case!(r#"
function sample(a) {}
const bound = sample.bind(null, 1, 2, 3);
console.log(bound.length);
"#, ["0"]);
}

#[test]
fn bind_name_has_bound_prefix() {
    case!(r#"
function sample() {}
const bound = sample.bind(null);
console.log(bound.name);
"#, ["bound sample"]);
}

#[test]
fn bind_bound_function_can_be_stored_in_object_and_invoked() {
    case!(r#"
function mult(a, b) {
    return a * b;
}
const obj = { fn: mult.bind(null, 6) };
console.log(obj.fn(7));
"#, ["42"]);
}

#[test]
fn bind_can_freeze_receiver_for_object_borrow() {
    case!(r#"
const borrow = Array.prototype.join.bind({ 0: "a", 1: "b", length: 2 }, "/");
console.log(borrow());
"#, ["a/b"]);
}

#[test]
fn bind_on_function_with_default_parameter_keeps_default_for_unbound_tail() {
    case!(r#"
function greet(greeting, name = "world") {
    return greeting + " " + name;
}
const hello = greet.bind(null, "hi");
console.log(hello());
"#, ["hi world"]);
}

#[test]
fn bind_on_function_with_rest_keeps_extra_args() {
    case!(r#"
function pack(head, ...rest) {
    return head + ":" + rest.join(",");
}
const fn = pack.bind(null, "a");
console.log(fn("b", "c"));
"#, ["a:b,c"]);
}

#[test]
fn arguments_length_matches_passed_values() {
    case!(r#"
function f(a) {
    console.log(arguments.length);
}
f(1, 2, 3);
"#, ["3"]);
}

#[test]
fn arguments_reads_extra_values_beyond_named_parameters() {
    case!(r#"
function f(a) {
    console.log(arguments[1]);
    console.log(arguments[2]);
}
f("x", "y", "z");
"#, ["y", "z"]);
}

#[test]
fn arguments_is_empty_when_no_values_passed() {
    case!(r#"
function f() {
    console.log(arguments.length);
    console.log(arguments[0]);
}
f();
"#, ["0", "undefined"]);
}

#[test]
fn arguments_write_updates_named_parameter_in_simple_list() {
    case!(r#"
function f(a) {
    arguments[0] = 7;
    console.log(a);
}
f(1);
"#, ["7"]);
}

#[test]
fn named_parameter_write_updates_arguments_in_simple_list() {
    case!(r#"
function f(a) {
    a = 8;
    console.log(arguments[0]);
}
f(1);
"#, ["8"]);
}

#[test]
fn arguments_object_can_be_joined_via_borrowed_array_method() {
    case!(r#"
function f() {
    console.log(Array.prototype.join.call(arguments, "-"));
}
f("a", "b", "c");
"#, ["a-b-c"]);
}

#[test]
fn arguments_inside_nested_arrow_uses_outer_call() {
    case!(r#"
function outer() {
    const inner = () => arguments[1];
    console.log(inner());
}
outer("a", "b");
"#, ["b"]);
}

#[test]
fn arguments_length_with_default_parameter_reflects_actual_call() {
    case!(r#"
function f(a = 1) {
    console.log(arguments.length);
}
f();
f(5);
"#, ["0", "1"]);
}

#[test]
fn arguments_with_default_parameter_does_not_alias_named_param() {
    case!(r#"
function f(a = 1) {
    arguments[0] = 9;
    console.log(a);
}
f(5);
"#, ["5"]);
}

#[test]
fn default_parameter_write_does_not_update_arguments_object() {
    case!(r#"
function f(a = 1) {
    a = 7;
    console.log(arguments[0]);
}
f(5);
"#, ["5"]);
}

#[test]
fn default_initializer_can_read_arguments_of_current_call() {
    case!(r#"
function f(a, b = arguments[0] + 1) {
    console.log(b);
}
f(4);
"#, ["5"]);
}

#[test]
fn rest_parameter_does_not_change_arguments_length() {
    case!(r#"
function f(a, ...rest) {
    console.log(arguments.length);
    console.log(rest.length);
}
f(1, 2, 3);
"#, ["3", "2"]);
}

#[test]
fn arguments_object_and_rest_collect_same_tail_values() {
    case!(r#"
function f(a, ...rest) {
    console.log(arguments[2]);
    console.log(rest[1]);
}
f("x", "y", "z");
"#, ["z", "z"]);
}

#[test]
fn arguments_object_survives_method_call_in_body() {
    case!(r#"
function f(a) {
    console.log(arguments[0]);
    console.log("x".toUpperCase());
    console.log(arguments[0]);
}
f("ok");
"#, ["ok", "X", "ok"]);
}

#[test]
fn default_parameter_uses_earlier_argument_only_when_tail_missing() {
    case!(r#"
function f(a, b = a + 1) {
    console.log(b);
}
f(2);
f(2, 10);
"#, ["3", "10"]);
}

#[test]
fn default_parameter_runs_on_explicit_undefined_not_null() {
    case!(r#"
function f(a = "x") {
    console.log(a);
}
f(undefined);
f(null);
"#, ["x", "null"]);
}

#[test]
fn rest_collects_no_extra_values_after_defaulted_param() {
    case!(r#"
function f(a = 1, ...rest) {
    console.log(a);
    console.log(rest.length);
}
f();
"#, ["1", "0"]);
}

#[test]
fn rest_collects_values_after_defaulted_head() {
    case!(r#"
function f(a = 1, ...rest) {
    console.log(a);
    console.log(rest.join(","));
}
f(undefined, 2, 3);
"#, ["1", "2,3"]);
}

#[test]
fn rest_array_can_be_applied_into_math_max() {
    case!(r#"
function f(...nums) {
    console.log(Math.max.apply(null, nums));
}
f(3, 8, 5);
"#, ["8"]);
}

#[test]
fn rest_array_can_be_spread_after_defaulted_prefix() {
    case!(r#"
function join(head = "x", ...tail) {
    console.log([head].concat(tail).join(","));
}
join(undefined, "a", "b");
"#, ["x,a,b"]);
}

#[test]
fn default_initializer_can_reference_bound_argument() {
    case!(r#"
function f(a, b = a * 2) {
    return b;
}
const g = f.bind(null, 4);
console.log(g());
"#, ["8"]);
}

#[test]
fn rest_parameter_length_is_zero_when_only_named_args_supplied() {
    case!(r#"
function f(a, b, ...rest) {
    console.log(rest.length);
}
f(1, 2);
"#, ["0"]);
}

#[test]
fn default_parameter_and_apply_can_forward_current_arguments() {
    case!(r#"
function wrap(a = 1) {
    function inner(x, y) {
        console.log(x + y);
    }
    inner.apply(null, arguments);
}
wrap(2, 3);
"#, ["5"]);
}

#[test]
fn arrow_inside_method_captures_receiver() {
    case!(r#"
const obj = {
    value: 2,
    run() {
        const f = () => this.value + 1;
        console.log(f());
    }
};
obj.run();
"#, ["3"]);
}

#[test]
fn arrow_call_cannot_rebind_this() {
    case!(r#"
const obj = {
    value: 2,
    run() {
        const f = () => this.value;
        console.log(f.call({ value: 9 }));
    }
};
obj.run();
"#, ["2"]);
}

#[test]
fn bound_arrow_ignores_bound_receiver() {
    case!(r#"
const make = function () {
    const arrow = () => this.name;
    return arrow.bind({ name: "bound" });
};
console.log(make.call({ name: "outer" })());
"#, ["outer"]);
}

#[test]
fn method_extraction_loses_receiver_but_call_restores_it() {
    case!(r#"
const obj = {
    name: "Ada",
    speak() {
        return this.name;
    }
};
const loose = obj.speak;
console.log(loose.call(obj));
"#, ["Ada"]);
}

#[test]
fn borrowed_object_method_operates_on_foreign_object() {
    case!(r#"
const source = {
    x: 10,
    get() {
        return this.x;
    }
};
const target = { x: 22 };
console.log(source.get.call(target));
"#, ["22"]);
}

#[test]
fn borrowed_array_map_operates_on_array_like_object() {
    case!(r#"
const arrLike = { 0: 1, 1: 2, length: 2 };
const out = Array.prototype.map.call(arrLike, x => x * 3);
console.log(out.join(","));
"#, ["3,6"]);
}

#[test]
fn borrowed_array_filter_operates_on_arguments_object() {
    case!(r#"
function f() {
    const out = Array.prototype.filter.call(arguments, x => x > 1);
    console.log(out.join(","));
}
f(1, 2, 3);
"#, ["2,3"]);
}

#[test]
fn borrowed_string_trim_can_process_primitive_via_call() {
    case!(r#"
console.log(String.prototype.trim.call("  hi  "));
"#, ["hi"]);
}

#[test]
fn function_prototype_call_can_invoke_borrowed_method_directly() {
    case!(r#"
const slice = String.prototype.slice;
console.log(Function.prototype.call.call(slice, "hello", 1, 4));
"#, ["ell"]);
}

#[test]
fn object_method_returning_arrow_survives_extraction() {
    case!(r#"
const obj = {
    value: 5,
    make() {
        return () => this.value;
    }
};
const maker = obj.make;
const arrow = maker.call(obj);
console.log(arrow());
"#, ["5"]);
}

#[test]
fn arrow_nested_inside_bound_function_reads_bound_receiver() {
    case!(r#"
function outer() {
    return () => this.label;
}
const fn = outer.bind({ label: "bound" })();
console.log(fn());
"#, ["bound"]);
}

#[test]
fn array_method_callback_can_be_bound_with_receiver() {
    case!(r#"
const ctx = { factor: 4 };
function mul(x) {
    return x * this.factor;
}
const out = [1, 2].map(mul.bind(ctx));
console.log(out.join(","));
"#, ["4,8"]);
}

#[test]
fn borrowed_join_bound_to_arguments_like_value_keeps_separator() {
    case!(r#"
function f() {
    const join = Array.prototype.join.bind(arguments, "/");
    console.log(join());
}
f("x", "y", "z");
"#, ["x/y/z"]);
}

#[test]
fn bound_constructor_uses_preset_leading_args() {
    case!(r#"
function Point(x, y) {
    this.x = x;
    this.y = y;
}
const PointY2 = Point.bind(null, 1);
const p = new PointY2(2);
console.log(p.x);
console.log(p.y);
"#, ["1", "2"]);
}

#[test]
fn new_on_bound_function_ignores_bound_this() {
    case!(r#"
function Person(name) {
    this.name = name;
}
const Bound = Person.bind({ name: "ignored" });
const p = new Bound("Ada");
console.log(p.name);
"#, ["Ada"]);
}

#[test]
fn constructed_bound_instance_is_instanceof_target() {
    case!(r#"
function Animal(kind) {
    this.kind = kind;
}
const Dog = Animal.bind(null, "dog");
const d = new Dog();
console.log(d instanceof Animal);
console.log(d.kind);
"#, ["true", "dog"]);
}

#[test]
fn bound_constructor_length_reflects_unbound_args() {
    case!(r#"
function Pair(a, b) {}
const One = Pair.bind(null, 1);
console.log(One.length);
"#, ["1"]);
}

#[test]
fn bound_constructor_name_has_prefix() {
    case!(r#"
function Pair(a, b) {}
const One = Pair.bind(null, 1);
console.log(One.name);
"#, ["bound Pair"]);
}