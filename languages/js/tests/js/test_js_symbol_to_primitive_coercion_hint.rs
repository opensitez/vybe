use super::helpers::run_js;

// ═══════════════════════════════════════════════════════════
// JavaScript: `Symbol.toPrimitive` Custom Coercion & Hint Handling
// ═══════════════════════════════════════════════════════════

#[test]
fn test_js_symbol_to_primitive_number_hint() {
    let src = r#"
const obj = {
    [Symbol.toPrimitive](hint) {
        if (hint === "number") return 42;
        return null;
    }
};
console.log(+obj + "|" + (obj * 2));
"#;
    assert_eq!(run_js(src), vec!["42|84"]);
}

#[test]
fn test_js_symbol_to_primitive_string_hint() {
    let src = r#"
const obj = {
    [Symbol.toPrimitive](hint) {
        if (hint === "string") return "CustomString";
        return null;
    }
};
console.log(String(obj));
"#;
    assert_eq!(run_js(src), vec!["CustomString"]);
}

#[test]
fn test_js_symbol_to_primitive_default_hint() {
    let src = r#"
const obj = {
    [Symbol.toPrimitive](hint) {
        if (hint === "default") return "DefaultValue";
        return "Other";
    }
};
console.log(obj + "!"); // '+' operator triggers 'default' hint
"#;
    assert_eq!(run_js(src), vec!["DefaultValue!"]);
}

#[test]
fn test_js_symbol_to_primitive_equality_comparison_default_hint() {
    let src = r#"
const obj = {
    [Symbol.toPrimitive](hint) {
        console.log("Hint: " + hint);
        return 100;
    }
};
console.log(obj == 100);
"#;
    assert_eq!(run_js(src), vec!["Hint: default", "true"]);
}

#[test]
fn test_js_symbol_to_primitive_overrides_valueOf_and_toString() {
    let src = r#"
const obj = {
    valueOf() { return 10; },
    toString() { return "10"; },
    [Symbol.toPrimitive](hint) {
        return 999;
    }
};
console.log(+obj + "|" + String(obj));
"#;
    assert_eq!(run_js(src), vec!["999|999"]);
}

#[test]
fn test_js_symbol_to_primitive_returning_object_throws_typeerror() {
    let src = r#"
const badObj = {
    [Symbol.toPrimitive]() {
        return {}; // Symbol.toPrimitive MUST return a primitive value!
    }
};
try {
    +badObj;
} catch (e) {
    console.log("ToPrimitive Non-Primitive TypeError");
}
"#;
    assert_eq!(run_js(src), vec!["ToPrimitive Non-Primitive TypeError"]);
}

#[test]
fn test_js_symbol_to_primitive_date_object_default_hint_is_string() {
    let src = r#"
const d = new Date(0);
console.log(typeof d[Symbol.toPrimitive] === "function");
"#;
    assert_eq!(run_js(src), vec!["true"]);
}

#[test]
fn test_js_symbol_to_primitive_bitwise_operators_number_hint() {
    let src = r#"
const obj = {
    [Symbol.toPrimitive](hint) {
        return hint === "number" ? 5 : 0;
    }
};
console.log((obj | 2) + "|" + (obj << 1));
"#;
    assert_eq!(run_js(src), vec!["7|10"]);
}

#[test]
fn test_js_symbol_to_primitive_template_literal_string_hint() {
    let src = r#"
const obj = {
    [Symbol.toPrimitive](hint) {
        return hint === "string" ? "TemplateStringVal" : "Wrong";
    }
};
console.log(`Value: ${obj}`);
"#;
    assert_eq!(run_js(src), vec!["Value: TemplateStringVal"]);
}

#[test]
fn test_js_symbol_to_primitive_relational_comparison_number_hint() {
    let src = r#"
const obj = {
    [Symbol.toPrimitive](hint) {
        return hint === "number" ? 15 : 0;
    }
};
console.log((obj > 10) + "|" + (obj < 20));
"#;
    assert_eq!(run_js(src), vec!["true|true"]);
}

#[test]
fn test_js_symbol_to_primitive_array_join_bypasses_to_primitive() {
    let src = r#"
const obj = {
    [Symbol.toPrimitive]() { return "Bypassed"; },
    toString() { return "CalledToString"; }
};
console.log([obj].join(""));
"#;
    assert_eq!(run_js(src), vec!["CalledToString"]);
}

#[test]
fn test_js_symbol_to_primitive_returning_symbol_primitive() {
    let src = r#"
const sym = Symbol("primitiveSym");
const obj = {
    [Symbol.toPrimitive]() { return sym; }
};
console.log(Object.is(obj[Symbol.toPrimitive](), sym));
"#;
    assert_eq!(run_js(src), vec!["true"]);
}

#[test]
fn test_js_symbol_to_primitive_returning_symbol_for_number_hint_throws_in_arithmetic() {
    let src = r#"
const obj = {
    [Symbol.toPrimitive]() { return Symbol("id"); }
};
try {
    +obj;
} catch (e) {
    console.log("Symbol Primitive in Number Conversion TypeError");
}
"#;
    assert_eq!(
        run_js(src),
        vec!["Symbol Primitive in Number Conversion TypeError"]
    );
}

#[test]
fn test_js_symbol_to_primitive_class_prototype_definition() {
    let src = r#"
class Money {
    constructor(amount, currency) {
        this.amount = amount;
        this.currency = currency;
    }
    [Symbol.toPrimitive](hint) {
        if (hint === "string") return `${this.amount} ${this.currency}`;
        return this.amount;
    }
}
const m = new Money(50, "USD");
console.log(String(m) + "|" + (m + 10));
"#;
    assert_eq!(run_js(src), vec!["50 USD|60"]);
}

#[test]
fn test_js_symbol_to_primitive_not_callable_throws_typeerror() {
    let src = r#"
const obj = { [Symbol.toPrimitive]: "not_a_function" };
try {
    +obj;
} catch (e) {
    console.log("ToPrimitive Not Callable TypeError");
}
"#;
    assert_eq!(run_js(src), vec!["ToPrimitive Not Callable TypeError"]);
}

#[test]
fn test_js_symbol_to_primitive_unary_minus_operator() {
    let src = r#"
const obj = {
    [Symbol.toPrimitive](hint) {
        return hint === "number" ? 25 : 0;
    }
};
console.log(-obj);
"#;
    assert_eq!(run_js(src), vec!["-25"]);
}

#[test]
fn test_js_symbol_to_primitive_property_key_computed_coercion() {
    let src = r#"
const keyObj = {
    [Symbol.toPrimitive](hint) {
        return hint === "string" ? "computedProp" : "wrong";
    }
};
const data = { [keyObj]: "TargetValue" };
console.log(data.computedProp);
"#;
    assert_eq!(run_js(src), vec!["TargetValue"]);
}

#[test]
fn test_js_symbol_to_primitive_bigint_coercion() {
    let src = r#"
const obj = {
    [Symbol.toPrimitive](hint) {
        return 100n;
    }
};
console.log(BigInt(obj).toString());
"#;
    assert_eq!(run_js(src), vec!["100"]);
}

#[test]
fn test_js_symbol_to_primitive_json_stringify_bypasses_to_primitive() {
    let src = r#"
const obj = {
    a: 1,
    [Symbol.toPrimitive]() { return "BypassedJSON"; }
};
console.log(JSON.stringify(obj));
"#;
    assert_eq!(run_js(src), vec![r#"{"a":1}"#]);
}

#[test]
fn test_js_symbol_to_primitive_well_known_symbol_identity() {
    let src = r#"
console.log(typeof Symbol.toPrimitive === "symbol");
"#;
    assert_eq!(run_js(src), vec!["true"]);
}

#[test]
fn test_js_symbol_to_primitive_returning_undefined_coerces_to_nan() {
    let src = r#"
const obj = { [Symbol.toPrimitive]() { return undefined; } };
console.log(Number.isNaN(+obj));
"#;
    assert_eq!(run_js(src), vec!["true"]);
}
