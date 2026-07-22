use super::helpers::run_js;

// ═══════════════════════════════════════════════════════════
// JavaScript: Object ToPrimitive Conversion Algorithm (`valueOf`, `toString`, `Symbol.toPrimitive`)
// ═══════════════════════════════════════════════════════════

#[test]
fn test_js_toprimitive_number_hint_prefers_valueof() {
    let src = r#"
const log = [];
const obj = {
    valueOf() { log.push("valueOf"); return 42; },
    toString() { log.push("toString"); return "42"; }
};
const res = Number(obj);
console.log(res + "|" + log.join(","));
"#;
    assert_eq!(run_js(src), vec!["42|valueOf"]); // "number" hint calls valueOf first!
}

#[test]
fn test_js_toprimitive_string_hint_prefers_tostring() {
    let src = r#"
const log = [];
const obj = {
    valueOf() { log.push("valueOf"); return 42; },
    toString() { log.push("toString"); return "hello"; }
};
const res = String(obj);
console.log(res + "|" + log.join(","));
"#;
    assert_eq!(run_js(src), vec!["hello|toString"]); // "string" hint calls toString first!
}

#[test]
fn test_js_toprimitive_fallback_when_preferred_returns_object() {
    let src = r#"
const log = [];
const obj = {
    valueOf() { log.push("valueOf"); return {}; }, // Returns object, not primitive!
    toString() { log.push("toString"); return "fallbackStr"; }
};
const res = Number(obj);
console.log(res + "|" + log.join(","));
"#;
    assert_eq!(run_js(src), vec!["NaN|valueOf,toString"]); // valueOf returns object, falls back to toString!
}

#[test]
fn test_js_toprimitive_symbol_to_primitive_overrides_both() {
    let src = r#"
const log = [];
const obj = {
    valueOf() { log.push("valueOf"); return 10; },
    toString() { log.push("toString"); return "str"; },
    [Symbol.toPrimitive](hint) { log.push("toPrimitive:" + hint); return 99; }
};
const res = obj + 1;
console.log(res + "|" + log.join(","));
"#;
    assert_eq!(run_js(src), vec!["100|toPrimitive:default"]);
}

#[test]
fn test_js_toprimitive_throws_typeerror_when_neither_returns_primitive() {
    let src = r#"
const obj = {
    valueOf() { return {}; },
    toString() { return {}; }
};
try {
    Number(obj);
} catch (e) {
    console.log("ToPrimitive TypeError");
}
"#;
    assert_eq!(run_js(src), vec!["ToPrimitive TypeError"]);
}

#[test]
fn test_js_toprimitive_symbol_to_primitive_must_return_primitive() {
    let src = r#"
const obj = {
    [Symbol.toPrimitive]() { return {}; }
};
try {
    +obj;
} catch (e) {
    console.log("Symbol.toPrimitive Non-Primitive TypeError");
}
"#;
    assert_eq!(
        run_js(src),
        vec!["Symbol.toPrimitive Non-Primitive TypeError"]
    );
}

#[test]
fn test_js_toprimitive_hints_in_various_operations() {
    let src = r#"
const log = [];
const obj = {
    [Symbol.toPrimitive](hint) {
        log.push(hint);
        return hint === "string" ? "str" : 10;
    }
};

String(obj); // string
Number(obj); // number
obj + 5;     // default
obj == 10;   // default
console.log(log.join(","));
"#;
    assert_eq!(run_js(src), vec!["string,number,default,default"]);
}

#[test]
fn test_js_toprimitive_date_object_default_hint_is_string() {
    let src = r#"
const log = [];
const d = new Date(0);
d[Symbol.toPrimitive] = function(hint) {
    log.push(hint);
    return Date.prototype[Symbol.toPrimitive].call(this, hint);
};
const res = d + 10;
console.log(log.join(",") + "|" + (typeof res));
"#;
    assert_eq!(run_js(src), vec!["default|string"]);
}

#[test]
fn test_js_valueof_inherited_from_object_prototype() {
    let src = r#"
const obj = {};
console.log(obj.valueOf() === obj); // Default Object.prototype.valueOf returns this!
"#;
    assert_eq!(run_js(src), vec!["true"]);
}

#[test]
fn test_js_tostring_inherited_from_object_prototype() {
    let src = r#"
const obj = {};
console.log(obj.toString()); // Default Object.prototype.toString returns [object Object]
"#;
    assert_eq!(run_js(src), vec!["[object Object]"]);
}

#[test]
fn test_js_toprimitive_null_prototype_object_requires_custom_methods() {
    let src = r#"
const obj = Object.create(null);
try {
    Number(obj);
} catch (e) {
    console.log("Null Prototype Object ToPrimitive TypeError");
}
"#;
    assert_eq!(
        run_js(src),
        vec!["Null Prototype Object ToPrimitive TypeError"]
    );
}

#[test]
fn test_js_toprimitive_null_prototype_object_with_valueof() {
    let src = r#"
const obj = Object.create(null);
obj.valueOf = () => 500;
console.log(Number(obj));
"#;
    assert_eq!(run_js(src), vec!["500"]);
}

#[test]
fn test_js_toprimitive_symbol_to_primitive_non_callable_throws_typeerror() {
    let src = r#"
const obj = {
    [Symbol.toPrimitive]: "not_a_function"
};
try {
    +obj;
} catch (e) {
    console.log("Symbol.toPrimitive Non-Callable TypeError");
}
"#;
    assert_eq!(
        run_js(src),
        vec!["Symbol.toPrimitive Non-Callable TypeError"]
    );
}

#[test]
fn test_js_toprimitive_valueof_non_callable_ignored() {
    let src = r#"
const obj = {
    valueOf: "not_a_function",
    toString: () => "validStr"
};
console.log(Number(obj)); // Non-callable valueOf is ignored, falls back to toString!
"#;
    assert_eq!(run_js(src), vec!["NaN"]);
}

#[test]
fn test_js_toprimitive_tostring_returning_number_in_string_hint() {
    let src = r#"
const obj = {
    toString() { return 123; } // toString can return any primitive!
};
console.log(String(obj));
"#;
    assert_eq!(run_js(src), vec!["123"]);
}

#[test]
fn test_js_toprimitive_valueof_returning_string_in_number_hint() {
    let src = r#"
const obj = {
    valueOf() { return "456"; } // valueOf returning string is converted to number by Number()!
};
console.log(Number(obj));
"#;
    assert_eq!(run_js(src), vec!["456"]);
}

#[test]
fn test_js_toprimitive_bigint_returned_by_symbol_to_primitive() {
    let src = r#"
const obj = {
    [Symbol.toPrimitive]() { return 100n; }
};
console.log((obj == 100n) + "|" + (typeof obj));
"#;
    assert_eq!(run_js(src), vec!["true|object"]);
}

#[test]
fn test_js_toprimitive_symbol_returned_by_symbol_to_primitive() {
    let src = r#"
const s = Symbol("mySym");
const obj = {
    [Symbol.toPrimitive]() { return s; }
};
console.log(String(obj) === "Symbol(mySym)");
"#;
    assert_eq!(run_js(src), vec!["true"]);
}

#[test]
fn test_js_toprimitive_this_binding_inside_valueof() {
    let src = r#"
const obj = {
    count: 7,
    valueOf() { return this.count * 2; }
};
console.log(+obj);
"#;
    assert_eq!(run_js(src), vec!["14"]);
}

#[test]
fn test_js_toprimitive_custom_class_instance() {
    let src = r#"
class Money {
    constructor(amount, currency) {
        this.amount = amount;
        this.currency = currency;
    }
    [Symbol.toPrimitive](hint) {
        if (hint === "number") return this.amount;
        if (hint === "string") return `${this.amount} ${this.currency}`;
        return this.amount;
    }
}
const m = new Money(50, "USD");
console.log(`${+m} | ${String(m)} | ${m + 10}`);
"#;
    assert_eq!(run_js(src), vec!["50 | 50 USD | 60"]);
}
