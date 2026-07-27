use super::helpers::run_js;

// ═══════════════════════════════════════════════════════════
// JavaScript: Logical Assignment Operators (`&&=`, `||=`, `??=`)
// ═══════════════════════════════════════════════════════════

#[test]
fn test_js_logical_and_assignment_truthy_target() {
    let src = r#"
let a = 1;
a &&= 2;
console.log(a);
"#;
    assert_eq!(run_js(src), vec!["2"]);
}

#[test]
fn test_js_logical_and_assignment_falsy_target_short_circuits() {
    let src = r#"
let a = 0;
let evaluated = false;
a &&= (evaluated = true, 5);
console.log(a + "|Evaluated=" + evaluated);
"#;
    assert_eq!(run_js(src), vec!["0|Evaluated=false"]); // Short-circuits: right side is not evaluated or assigned!
}

#[test]
fn test_js_logical_or_assignment_falsy_target() {
    let src = r#"
let a = 0;
a ||= 10;
console.log(a);
"#;
    assert_eq!(run_js(src), vec!["10"]);
}

#[test]
fn test_js_logical_or_assignment_truthy_target_short_circuits() {
    let src = r#"
let a = 5;
let evaluated = false;
a ||= (evaluated = true, 10);
console.log(a + "|Evaluated=" + evaluated);
"#;
    assert_eq!(run_js(src), vec!["5|Evaluated=false"]);
}

#[test]
fn test_js_nullish_assignment_null_target() {
    let src = r#"
let a = null;
a ??= "default";
console.log(a);
"#;
    assert_eq!(run_js(src), vec!["default"]);
}

#[test]
fn test_js_nullish_assignment_undefined_target() {
    let src = r#"
let a = undefined;
a ??= "fallback";
console.log(a);
"#;
    assert_eq!(run_js(src), vec!["fallback"]);
}

#[test]
fn test_js_nullish_assignment_falsy_non_nullish_target_short_circuits() {
    let src = r#"
let a = 0;
let b = "";
let c = false;

a ??= 99;
b ??= "default";
c ??= true;
console.log(`${a}:${b}:${c}`);
"#;
    assert_eq!(run_js(src), vec!["0::false"]); // 0, "", false are not nullish, so ??= does NOT assign!
}

#[test]
fn test_js_logical_assignment_object_property_accessors() {
    let src = r#"
const obj = { x: null, y: "exist" };
obj.x ??= "defaultX";
obj.y ??= "defaultY";
console.log(`${obj.x}:${obj.y}`);
"#;
    assert_eq!(run_js(src), vec!["defaultX:exist"]);
}

#[test]
fn test_js_logical_assignment_no_setter_call_when_short_circuited() {
    let src = r#"
let setterCount = 0;
const obj = {
    _val: "Initial",
    get val() { return this._val; },
    set val(v) { setterCount++; this._val = v; }
};
obj.val ||= "NewValue"; // Initial is truthy -> short circuits, setter NOT called!
console.log(obj.val + "|Setters=" + setterCount);
"#;
    assert_eq!(run_js(src), vec!["Initial|Setters=0"]);
}

#[test]
fn test_js_logical_assignment_and_assignment_short_circuited_by_falsy_property() {
    let src = r#"
let setterCalls = 0;
let rhsExecuted = false;

const obj = {
    _x: 0,
    get x() {
        return this._x;
    },
    set x(v) {
        setterCalls++;
        this._x = v;
    },
};

    obj.x &&= (rhsExecuted = true, 99);
console.log(`${obj.x}|${setterCalls}|${rhsExecuted}`);

obj._x = 1;
rhsExecuted = false;
obj.x &&= (rhsExecuted = true, 33);
console.log(`${obj.x}|${setterCalls}|${rhsExecuted}`);
"#;
    assert_eq!(run_js(src), vec!["0|1|false", "33|2|true"]);
}

#[test]
fn test_js_logical_assignment_setter_called_when_evaluated() {
    let src = r#"
let setterCount = 0;
const obj = {
    _val: null,
    get val() { return this._val; },
    set val(v) { setterCount++; this._val = v; }
};
obj.val ??= "AssignedValue";
console.log(obj.val + "|Setters=" + setterCount);
"#;
    assert_eq!(run_js(src), vec!["AssignedValue|Setters=1"]);
}

#[test]
fn test_js_logical_assignment_array_element() {
    let src = r#"
const arr = [0, null, 10];
arr[0] ||= 1;
arr[1] ??= 2;
arr[2] &&= 30;
console.log(arr.join(","));
"#;
    assert_eq!(run_js(src), vec!["1,2,30"]);
}

#[test]
fn test_js_logical_assignment_computed_property_evaluated_once() {
    let src = r#"
let keyEvaluationCount = 0;
const getKey = () => { keyEvaluationCount++; return "prop"; };
const obj = { prop: null };

obj[getKey()] ??= "Assigned";
console.log(obj.prop + "|KeyEvalCount=" + keyEvaluationCount);
"#;
    assert_eq!(run_js(src), vec!["Assigned|KeyEvalCount=1"]);
}

#[test]
fn test_js_logical_assignment_in_function_parameters() {
    let src = r#"
function fn(opts) {
    opts ||= {};
    opts.timeout ??= 1000;
    return opts.timeout;
}
console.log(fn() + "|" + fn({ timeout: 500 }));
"#;
    assert_eq!(run_js(src), vec!["1000|500"]);
}

#[test]
fn test_js_logical_assignment_chained() {
    let src = r#"
let x = null;
x ??= 0;
x ||= 10;
x &&= 20;
console.log(x);
"#;
    assert_eq!(run_js(src), vec!["20"]);
}

#[test]
fn test_js_logical_assignment_const_reassignment_throws_typeerror() {
    let src = r#"
const x = 0;
try {
    eval("x ||= 10;");
} catch (e) {
    console.log("Const Logical Assignment TypeError");
}
"#;
    assert_eq!(run_js(src), vec!["Const Logical Assignment TypeError"]);
}

#[test]
fn test_js_logical_assignment_const_short_circuited_still_throws_syntaxerror_or_typeerror() {
    let src = r#"
const x = 5;
try {
    eval("x ||= 10;");
} catch (e) {
    console.log("Const Logical Assignment Reassignment Error");
}
"#;
    assert_eq!(
        run_js(src),
        vec!["Const Logical Assignment Reassignment Error"]
    );
}

#[test]
fn test_js_logical_assignment_class_private_field() {
    let src = r#"
class Cache {
    #data = null;
    get() {
        return (this.#data ??= "CachedData");
    }
}
const c = new Cache();
console.log(c.get() + "|" + c.get());
"#;
    assert_eq!(run_js(src), vec!["CachedData|CachedData"]);
}

#[test]
fn test_js_logical_assignment_eval_return_value() {
    let src = r#"
let a = null;
console.log(eval("a ??= 99;"));
"#;
    assert_eq!(run_js(src), vec!["99"]);
}

#[test]
fn test_js_logical_assignment_bigint_and_assignment() {
    let src = r#"
let b = 10n;
b &&= 20n;
console.log(b.toString());
"#;
    assert_eq!(run_js(src), vec!["20"]);
}

#[test]
fn test_js_logical_assignment_symbol_or_assignment() {
    let src = r#"
let sym = null;
const targetSym = Symbol("assigned");
sym ||= targetSym;
console.log(sym === targetSym);
"#;
    assert_eq!(run_js(src), vec!["true"]);
}

#[test]
fn test_js_logical_assignment_accessor_setter_getter_calls() {
    let src = r#"
let getCount = 0;
let setCount = 0;
const obj = {
    _value: null,
    get value() {
        getCount++;
        return this._value;
    },
    set value(v) {
        setCount++;
        this._value = v;
    }
};

obj.value ||= "fallback";
obj.value ||= "ignored";
obj.value &&= "updated";

console.log(`${obj.value}|${getCount}|${setCount}`);
"#;
assert_eq!(run_js(src), vec!["updated|4|3"]);
}

#[test]
fn test_js_logical_assignment_computed_property_uses_rhs_short_circuit_per_operator() {
    let src = r#"
let keyEval = 0;
const key = () => {
    keyEval++;
    return "value";
};

const obj = {
    value: null,
};

obj[key()] ||= "filled"; // assigns
obj[key()] ||= "ignored"; // short-circuit, no assign
obj[key()] &&= "final";  // assigns

console.log(obj.value);
console.log(keyEval);
"#;

    assert_eq!(
        run_js(src),
        vec!["final", "6"]
    );
}
