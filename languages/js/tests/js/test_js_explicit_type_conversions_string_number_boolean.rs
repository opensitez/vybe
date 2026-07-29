use super::helpers::run_js;

// ═══════════════════════════════════════════════════════════
// JavaScript: Explicit Type Conversions (`String()`, `Number()`, `Boolean()`, `BigInt()`)
// ═══════════════════════════════════════════════════════════

#[test]
fn test_js_string_constructor_primitives() {
    let src = r#"
console.log([
    String(123),
    String(true),
    String(false),
    String(null),
    String(undefined),
    String(10n),
    String(Symbol("id"))
].join("|"));
"#;
    assert_eq!(
        run_js(src),
        vec!["123|true|false|null|undefined|10|Symbol(id)"]
    );
}

#[test]
fn test_js_number_constructor_primitives() {
    let src = r#"
console.log([
    Number("42"),
    Number("  3.14  "),
    Number(""),
    Number("   "),
    Number(true),
    Number(false),
    Number(null),
    Number(undefined),
    Number("invalid")
].join("|"));
"#;
    assert_eq!(run_js(src), vec!["42|3.14|0|0|1|0|0|NaN|NaN"]);
}

#[test]
fn test_js_boolean_constructor_falsy_values() {
    let src = r#"
console.log([
    Boolean(false),
    Boolean(0),
    Boolean(-0),
    Boolean(0n),
    Boolean(""),
    Boolean(null),
    Boolean(undefined),
    Boolean(NaN)
].every(val => val === false));
"#;
    assert_eq!(run_js(src), vec!["true"]);
}

#[test]
fn test_js_boolean_constructor_truthy_values() {
    let src = r#"
console.log([
    Boolean(true),
    Boolean(1),
    Boolean(-1),
    Boolean("0"),
    Boolean("false"),
    Boolean([]),
    Boolean({}),
    Boolean(() => {})
].every(val => val === true));
"#;
    assert_eq!(run_js(src), vec!["true"]);
}

#[test]
fn test_js_bigint_constructor_conversions() {
    let src = r#"
console.log([
    BigInt(100).toString(),
    BigInt("200").toString(),
    BigInt("0b1010").toString(),
    BigInt("0xff").toString(),
    BigInt(true).toString(),
    BigInt(false).toString()
].join("|"));
"#;
    assert_eq!(run_js(src), vec!["100|200|10|255|1|0"]);
}

#[test]
fn test_js_bigint_constructor_float_throws_rangeerror() {
    let src = r#"
try {
    BigInt(3.14);
} catch (e) {
    console.log("BigInt Float RangeError");
}
"#;
    assert_eq!(run_js(src), vec!["BigInt Float RangeError"]);
}

#[test]
fn test_js_bigint_constructor_null_or_undefined_throws_typeerror() {
    let src = r#"
try {
    BigInt(null);
} catch (e) {
    console.log("BigInt Null TypeError");
}
"#;
    assert_eq!(run_js(src), vec!["BigInt Null TypeError"]);
}

#[test]
fn test_js_parse_int_radix_conversions() {
    let src = r#"
console.log([
    parseInt("42", 10),
    parseInt("1010", 2),
    parseInt("ff", 16),
    parseInt("077", 8),
    parseInt("100px", 10),
    parseInt("abc", 10)
].join("|"));
"#;
    assert_eq!(run_js(src), vec!["42|10|255|77|100|NaN"]);
}

#[test]
fn test_js_parse_float_conversions() {
    let src = r#"
console.log([
    parseFloat("3.14"),
    parseFloat("314e-2"),
    parseFloat("10.5.6"),
    parseFloat("  42  "),
    parseFloat("text")
].join("|"));
"#;
    assert_eq!(run_js(src), vec!["3.14|3.14|10.5|42|NaN"]);
}

#[test]
fn test_js_number_constructor_symbol_throws_typeerror() {
    let src = r#"
try {
    Number(Symbol("sym"));
} catch (e) {
    console.log("Number Symbol TypeError");
}
"#;
    assert_eq!(run_js(src), vec!["Number Symbol TypeError"]);
}

#[test]
fn test_js_string_constructor_objects() {
    let src = r#"
console.log([
    String([1, 2, 3]),
    String({}),
    String({ toString: () => "customStr" })
].join("|"));
"#;
    assert_eq!(run_js(src), vec!["1,2,3|[object Object]|customStr"]);
}

#[test]
fn test_js_number_constructor_objects() {
    let src = r#"
console.log([
    Number([42]),
    Number([]),
    Number([1, 2]),
    Number({ valueOf: () => 99 })
].join("|"));
"#;
    assert_eq!(run_js(src), vec!["42|0|NaN|99"]);
}

#[test]
fn test_js_number_to_fixed_formatting() {
    let src = r#"
console.log([
    (123.456).toFixed(2),
    (123.4).toFixed(3),
    (0).toFixed(1)
].join("|"));
"#;
    assert_eq!(run_js(src), vec!["123.46|123.400|0.0"]);
}

#[test]
fn test_js_number_to_precision_formatting() {
    let src = r#"
console.log([
    (123.456).toPrecision(4),
    (0.00123).toPrecision(2)
].join("|"));
"#;
    assert_eq!(run_js(src), vec!["123.5|0.0012"]);
}

#[test]
fn test_js_number_to_exponential_formatting() {
    let src = r#"
console.log([
    (123456).toExponential(2),
    (0.005).toExponential(1)
].join("|"));
"#;
    assert_eq!(run_js(src), vec!["1.23e+5|5.0e-3"]);
}

#[test]
fn test_js_number_to_string_radix() {
    let src = r#"
console.log([
    (255).toString(16),
    (10).toString(2),
    (64).toString(8)
].join("|"));
"#;
    assert_eq!(run_js(src), vec!["ff|1010|100"]);
}

#[test]
fn test_js_bigint_to_string_radix() {
    let src = r#"
console.log([
    (255n).toString(16),
    (10n).toString(2)
].join("|"));
"#;
    assert_eq!(run_js(src), vec!["ff|1010"]);
}

#[test]
fn test_js_parse_int_radix_bounds() {
    let src = r#"
try {
    parseInt("10", 37); // Radix must be between 2 and 36!
} catch (e) {
    console.log("parseInt Invalid Radix");
}
console.log(isNaN(parseInt("10", 37)));
"#;
    assert_eq!(run_js(src), vec!["true"]);
}

#[test]
fn test_js_explicit_type_conversions_without_new_vs_with_new() {
    let src = r#"
const primStr = String("hello");
const objStr = new String("hello");
console.log(`${typeof primStr}:${typeof objStr}:${primStr === objStr}`);
"#;
    assert_eq!(run_js(src), vec!["string:object:false"]);
}

#[test]
fn test_js_symbol_to_string_explicit() {
    let src = r#"
const s = Symbol("desc");
console.log(String(s) + "|" + s.toString());
"#;
    assert_eq!(run_js(src), vec!["Symbol(desc)|Symbol(desc)"]);
}

#[test]
fn test_js_bigint_constructor_symbol_throws_typeerror() {
    let src = r#"
try {
    BigInt(Symbol("foo"));
} catch (e) {
    console.log(e.name);
}
"#;
    assert_eq!(run_js(src), vec!["TypeError"]);
}

