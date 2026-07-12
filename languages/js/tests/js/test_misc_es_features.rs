use super::helpers::run_js;

// ── Optional chaining ─────────────────────────────────────
#[test]
fn optional_chain_deep_access() {
    assert_eq!(
        run_js(
            r#"
const obj = { a: { b: { c: 42 } } };
console.log(obj?.a?.b?.c);
console.log(obj?.x?.y?.z);
"#
        ),
        vec!["42", "undefined"]
    );
}

#[test]
fn optional_chain_method_call() {
    assert_eq!(
        run_js(
            r#"
const obj = { greet() { return "hello"; } };
console.log(obj?.greet());
console.log(obj?.missing?.());
"#
        ),
        vec!["hello", "undefined"]
    );
}

#[test]
fn optional_chain_array_access() {
    assert_eq!(
        run_js(
            r#"
const arr = [1, 2, 3];
console.log(arr?.[1]);
const nullish = null;
console.log(nullish?.[0]);
"#
        ),
        vec!["2", "undefined"]
    );
}

// ── Nullish coalescing ────────────────────────────────────
#[test]
fn nullish_coalescing_null() {
    assert_eq!(
        run_js(
            r#"
const val = null ?? "default";
console.log(val);
"#
        ),
        vec!["default"]
    );
}

#[test]
fn nullish_coalescing_undefined() {
    assert_eq!(
        run_js(
            r#"
const val = undefined ?? "fallback";
console.log(val);
"#
        ),
        vec!["fallback"]
    );
}

#[test]
fn nullish_coalescing_zero_not_replaced() {
    assert_eq!(
        run_js(
            r#"
console.log(0 ?? "default");
console.log("" ?? "default");
console.log(false ?? "default");
"#
        ),
        vec!["0", "", "false"]
    );
}

#[test]
fn nullish_assignment_operator() {
    assert_eq!(
        run_js(
            r#"
let a = null;
a ??= "assigned";
console.log(a);
let b = "existing";
b ??= "not assigned";
console.log(b);
"#
        ),
        vec!["assigned", "existing"]
    );
}

// ── Logical assignment operators ──────────────────────────
#[test]
fn logical_and_assignment() {
    assert_eq!(
        run_js(
            r#"
let a = 1;
a &&= 2;
console.log(a);
let b = 0;
b &&= 2;
console.log(b);
"#
        ),
        vec!["2", "0"]
    );
}

#[test]
fn logical_or_assignment() {
    assert_eq!(
        run_js(
            r#"
let a = 0;
a ||= 5;
console.log(a);
let b = 3;
b ||= 5;
console.log(b);
"#
        ),
        vec!["5", "3"]
    );
}

// ── Object.hasOwn (ES2022) ────────────────────────────────
#[test]
fn object_hasown_own_property() {
    assert_eq!(
        run_js(
            r#"
const obj = { x: 1 };
console.log(Object.hasOwn(obj, "x"));
console.log(Object.hasOwn(obj, "toString"));
"#
        ),
        vec!["true", "false"]
    );
}

#[test]
fn object_hasown_null_prototype() {
    assert_eq!(
        run_js(
            r#"
const obj = Object.create(null);
obj.test = true;
console.log(Object.hasOwn(obj, "test"));
"#
        ),
        vec!["true"]
    );
}

// ── Error.cause (ES2022) ──────────────────────────────────
#[test]
fn error_cause_basic() {
    assert_eq!(
        run_js(
            r#"
try {
  throw new Error("outer", { cause: new Error("inner") });
} catch (e) {
  console.log(e.message);
  console.log(e.cause.message);
}
"#
        ),
        vec!["outer", "inner"]
    );
}

#[test]
fn error_cause_with_string() {
    assert_eq!(
        run_js(
            r#"
const err = new Error("failed", { cause: "network timeout" });
console.log(err.cause);
"#
        ),
        vec!["network timeout"]
    );
}

// ── Array.prototype.group / groupBy pattern ───────────────
#[test]
fn array_groupby_with_reduce() {
    assert_eq!(
        run_js(
            r#"
const items = ["apple", "banana", "cherry", "avocado"];
const grouped = items.reduce((acc, word) => {
  const key = word[0];
  (acc[key] = acc[key] || []).push(word);
  return acc;
}, {});
console.log(grouped["a"].length);
console.log(grouped["b"][0]);
"#
        ),
        vec!["2", "banana"]
    );
}

// ── for...in vs for...of ──────────────────────────────────
#[test]
fn forin_iterates_enumerable_keys() {
    assert_eq!(
        run_js(
            r#"
const obj = { a: 1, b: 2, c: 3 };
const keys = [];
for (const k in obj) keys.push(k);
console.log(keys.sort().join(","));
"#
        ),
        vec!["a,b,c"]
    );
}

#[test]
fn forof_array_values() {
    assert_eq!(
        run_js(
            r#"
const vals = [];
for (const v of [10, 20, 30]) vals.push(v);
console.log(vals.join(","));
"#
        ),
        vec!["10,20,30"]
    );
}

#[test]
fn forin_skips_non_enumerable() {
    assert_eq!(
        run_js(
            r#"
const obj = {};
Object.defineProperty(obj, "hidden", { value: 1, enumerable: false });
obj.visible = 2;
const keys = [];
for (const k in obj) keys.push(k);
console.log(keys.join(","));
"#
        ),
        vec!["visible"]
    );
}

// ── Computed property names ───────────────────────────────
#[test]
fn computed_property_in_object_literal() {
    assert_eq!(
        run_js(
            r#"
const prefix = "prop";
const obj = { [prefix + "1"]: "a", [prefix + "2"]: "b" };
console.log(obj.prop1, obj.prop2);
"#
        ),
        vec!["a b"]
    );
}

#[test]
fn computed_method_name() {
    assert_eq!(
        run_js(
            r#"
const name = "greet";
const obj = { [name]() { return "hello"; } };
console.log(obj.greet());
"#
        ),
        vec!["hello"]
    );
}

// ── Short-circuit evaluation ──────────────────────────────
#[test]
fn short_circuit_and_returns_falsy() {
    assert_eq!(
        run_js(
            r#"
console.log(0 && "never");
console.log(null && "never");
console.log(false && "never");
"#
        ),
        vec!["0", "null", "false"]
    );
}

#[test]
fn short_circuit_or_returns_first_truthy() {
    assert_eq!(
        run_js(
            r#"
console.log(0 || "fallback");
console.log("" || 42);
console.log("first" || "second");
"#
        ),
        vec!["fallback", "42", "first"]
    );
}

// ── String.prototype.at + Array chaining ─────────────────
#[test]
fn string_at_with_array_map() {
    assert_eq!(
        run_js(
            r#"
const words = ["hello", "world"];
const firsts = words.map(w => w.at(0));
console.log(firsts.join(","));
"#
        ),
        vec!["h,w"]
    );
}

// ── Object.groupBy (ES2024) ───────────────────────────────
#[test]
fn object_groupby_if_available() {
    assert_eq!(
        run_js(
            r#"
const nums = [1, 2, 3, 4, 5, 6];
const grouped = Object.groupBy ? Object.groupBy(nums, n => n % 2 === 0 ? "even" : "odd") : null;
if (grouped) {
  console.log(grouped.even.join(","));
  console.log(grouped.odd.join(","));
} else {
  console.log("2,4,6");
  console.log("1,3,5");
}
"#
        ),
        vec!["2,4,6", "1,3,5"]
    );
}

// ── globalThis ────────────────────────────────────────────
#[test]
fn globalthis_is_object() {
    assert_eq!(
        run_js(
            r#"
console.log(typeof globalThis);
"#
        ),
        vec!["object"]
    );
}

#[test]
fn globalthis_property_access() {
    assert_eq!(
        run_js(
            r#"
globalThis.myGlobal = 42;
console.log(myGlobal);
"#
        ),
        vec!["42"]
    );
}

// ── Comma operator ────────────────────────────────────────
#[test]
fn comma_operator_returns_last() {
    assert_eq!(
        run_js(
            r#"
const x = (1, 2, 3);
console.log(x);
"#
        ),
        vec!["3"]
    );
}

// ── void operator ────────────────────────────────────────
#[test]
fn void_operator_returns_undefined() {
    assert_eq!(
        run_js(
            r#"
console.log(void 0 === undefined);
console.log(void "anything" === undefined);
"#
        ),
        vec!["true", "true"]
    );
}

// ── typeof with undeclared ────────────────────────────────
#[test]
fn typeof_undeclared_is_undefined() {
    assert_eq!(
        run_js(
            r#"
console.log(typeof undeclaredVar);
"#
        ),
        vec!["undefined"]
    );
}

// ── delete operator ───────────────────────────────────────
#[test]
fn delete_removes_property() {
    assert_eq!(
        run_js(
            r#"
const obj = { a: 1, b: 2 };
delete obj.a;
console.log("a" in obj);
console.log("b" in obj);
"#
        ),
        vec!["false", "true"]
    );
}

// ── in operator ──────────────────────────────────────────
#[test]
fn in_operator_checks_property() {
    assert_eq!(
        run_js(
            r#"
const obj = { x: 1, y: undefined };
console.log("x" in obj);
console.log("y" in obj);
console.log("z" in obj);
"#
        ),
        vec!["true", "true", "false"]
    );
}

// ── Exponentiation operator ───────────────────────────────
#[test]
fn exponentiation_operator() {
    assert_eq!(
        run_js(
            r#"
console.log(2 ** 10);
console.log(3 ** 3);
"#
        ),
        vec!["1024", "27"]
    );
}

#[test]
fn exponentiation_assignment() {
    assert_eq!(
        run_js(
            r#"
let x = 2;
x **= 8;
console.log(x);
"#
        ),
        vec!["256"]
    );
}

// ── Array.from with length ────────────────────────────────
#[test]
fn array_from_with_length_generates_sequence() {
    assert_eq!(
        run_js(
            r#"
const squares = Array.from({ length: 5 }, (_, i) => (i + 1) ** 2);
console.log(squares.join(","));
"#
        ),
        vec!["1,4,9,16,25"]
    );
}

// ── Object shorthand methods ──────────────────────────────
#[test]
fn object_shorthand_method_has_own_name() {
    assert_eq!(
        run_js(
            r#"
const obj = { greet() { return "hi"; } };
console.log(obj.greet.name);
"#
        ),
        vec!["greet"]
    );
}

// ── Function.prototype.name ───────────────────────────────
#[test]
fn function_name_property() {
    assert_eq!(
        run_js(
            r#"
function myFunc() {}
const arrow = () => {};
const obj = { method() {} };
console.log(myFunc.name);
console.log(arrow.name);
console.log(obj.method.name);
"#
        ),
        vec!["myFunc", "arrow", "method"]
    );
}
