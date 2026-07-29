/// Coercion edge cases — ToPrimitive, valueOf, toString, type juggling
use super::helpers::run_js;

#[test]
fn valueof_in_arithmetic() {
    assert_eq!(
        run_js(
            r#"
const obj = { valueOf() { return 42; } };
console.log(obj + 8);
console.log(obj * 2);
console.log(obj - 10);
"#
        ),
        vec!["50", "84", "32"]
    );
}

#[test]
fn tostring_in_string_context() {
    assert_eq!(
        run_js(
            r#"
const obj = { toString() { return "hello"; } };
console.log("" + obj);
console.log(`${obj}`);
"#
        ),
        vec!["hello", "hello"]
    );
}

#[test]
fn toprimitive_overrides_both() {
    assert_eq!(
        run_js(
            r#"
const obj = {
    [Symbol.toPrimitive](hint) {
        switch (hint) {
            case "number": return 42;
            case "string": return "forty-two";
            default: return true;
        }
    }
};
console.log(+obj);
console.log(`${obj}`);
console.log(obj + "");
"#
        ),
        vec!["42", "forty-two", "true"]
    );
}

#[test]
fn double_equals_coercion_table() {
    assert_eq!(
        run_js(
            r#"
// Key == behaviors
console.log(null == undefined);
console.log(null == 0);
console.log(0 == false);
console.log("" == false);
console.log("1" == 1);
console.log([] == false);
"#
        ),
        vec!["true", "false", "true", "true", "true", "true"]
    );
}

#[test]
fn array_plus_behaviors() {
    assert_eq!(
        run_js(
            r#"
console.log([] + []);
console.log([] + {});
console.log({} + []);
console.log([1] + [2]);
"#
        ),
        vec!["", "[object Object]", "[object Object]", "12"]
    );
}

#[test]
fn string_to_number_edge_cases() {
    assert_eq!(
        run_js(
            r#"
console.log(Number(" 42 ")); // trims
console.log(Number("0x10")); // hex
console.log(Number("0o10")); // octal
console.log(Number("0b10")); // binary
console.log(Number("Infinity"));
console.log(Number(""));
"#
        ),
        vec!["42", "16", "8", "2", "Infinity", "0"]
    );
}

#[test]
fn boolean_coercion_falsy_values() {
    assert_eq!(
        run_js(
            r#"
const falsies = [false, 0, "", null, undefined, NaN, -0, 0n];
console.log(falsies.every(v => !v));
const truthy = [1, "a", {}, [], () => {}, Infinity];
console.log(truthy.every(v => !!v));
"#
        ),
        vec!["true", "true"]
    );
}

#[test]
fn object_to_boolean_always_truthy() {
    assert_eq!(
        run_js(
            r#"
// All objects (even empty) are truthy
console.log(!!{});
console.log(!![]);
console.log(!!new Boolean(false));
console.log(!!new Number(0));
"#
        ),
        vec!["true", "true", "true", "true"]
    );
}

#[test]
fn comparison_object_to_primitive() {
    assert_eq!(
        run_js(
            r#"
// Date uses valueOf (timestamp) for comparisons
const d1 = new Date(0);
const d2 = new Date(1000);
console.log(d1 < d2);
console.log(d2 - d1);
"#
        ),
        vec!["true", "1000"]
    );
}

#[test]
fn unary_plus_on_objects() {
    assert_eq!(
        run_js(
            r#"
console.log(+null);
console.log(+undefined);
console.log(+true);
console.log(+false);
console.log(+[]);
console.log(+[1]);
console.log(+[1,2]);
"#
        ),
        vec!["0", "NaN", "1", "0", "0", "1", "NaN"]
    );
}

#[test]
fn valueof_tostring_are_tied_to_primitive_hints() {
    assert_eq!(
        run_js(
            r#"
const trace = [];
const obj = {
    valueOf() {
        trace.push("valueOf");
        return 4;
    },
    toString() {
        trace.push("toString");
        return "9";
    }
};
console.log(+obj);
console.log(`${obj}`);
console.log(trace.join(","));
"#
        ),
        vec!["4", "9", "valueOf,toString"]
    );
}

#[test]
fn object_to_primitive_must_return_primitive() {
    assert_eq!(
        run_js(
            r#"
const bad = {
    valueOf() {
        throw new Error("valueOf boom");
    },
    toString() {
        return "ok";
    }
};
try {
    console.log(+bad);
} catch (e) {
    console.log(e.message);
}
console.log(String(bad));
"#
        ),
        vec!["valueOf boom", "ok"]
    );
}

#[test]
fn test_toprimitive_returning_object_throws_typeerror() {
    assert_eq!(
        run_js(
            r#"
const bad = {
    [Symbol.toPrimitive]() {
        return {};
    }
};
try {
    +bad;
} catch (e) {
    console.log(e.name);
}
"#
        ),
        vec!["TypeError"]
    );
}

