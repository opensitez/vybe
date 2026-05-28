/// Numeric type coercions — Number(), parseInt(), parseFloat(), edge cases

use super::helpers::run_js;

#[test]
fn number_from_various_types() {
    assert_eq!(run_js(r#"
console.log(Number("42"));
console.log(Number("  3.14  ")); // trims whitespace
console.log(Number(""));
console.log(Number(null));
console.log(Number(undefined));
console.log(Number(true));
console.log(Number(false));
"#), vec!["42", "3.14", "0", "0", "NaN", "1", "0"]);
}

#[test]
fn number_from_array() {
    assert_eq!(run_js(r#"
console.log(Number([]));
console.log(Number([42]));
console.log(Number([1, 2]));
"#), vec!["0", "42", "NaN"]);
}

#[test]
fn parse_int_radix() {
    assert_eq!(run_js(r#"
console.log(parseInt("0xFF", 16));
console.log(parseInt("1010", 2));
console.log(parseInt("37", 8));
console.log(parseInt("z", 36));
"#), vec!["255", "10", "31", "35"]);
}

#[test]
fn parse_int_stops_at_invalid_char() {
    assert_eq!(run_js(r#"
console.log(parseInt("123abc"));
console.log(parseInt("0x1G", 16));
"#), vec!["123", "1"]);
}

#[test]
fn parse_int_leading_whitespace_ok() {
    assert_eq!(run_js(r#"
console.log(parseInt("  42  "));
"#), vec!["42"]);
}

#[test]
fn parse_float_basic() {
    assert_eq!(run_js(r#"
console.log(parseFloat("3.14"));
console.log(parseFloat(".5"));
console.log(parseFloat("1e3"));
console.log(parseFloat("Infinity"));
"#), vec!["3.14", "0.5", "1000", "Infinity"]);
}

#[test]
fn parse_float_stops_at_second_dot() {
    assert_eq!(run_js(r#"
console.log(parseFloat("3.14.15"));
"#), vec!["3.14"]);
}

#[test]
fn number_hex_octal_binary_literals() {
    assert_eq!(run_js(r#"
console.log(0xFF);
console.log(0o77);
console.log(0b1010);
"#), vec!["255", "63", "10"]);
}

#[test]
fn number_to_fixed_rounding() {
    assert_eq!(run_js(r#"
console.log((1.005).toFixed(2));
console.log((1.255).toFixed(2));
console.log((1.5).toFixed(0));
"#), vec!["1.00", "1.25", "2"]);
}

#[test]
fn number_to_string_radix() {
    assert_eq!(run_js(r#"
console.log((255).toString(16));
console.log((10).toString(2));
console.log((31).toString(8));
"#), vec!["ff", "1010", "37"]);
}

#[test]
fn unary_plus_coerces() {
    assert_eq!(run_js(r#"
console.log(+"42");
console.log(+true);
console.log(+false);
console.log(+null);
console.log(+undefined);
console.log(+"");
"#), vec!["42", "1", "0", "0", "NaN", "0"]);
}

#[test]
fn bitwise_coerces_to_int32() {
    assert_eq!(run_js(r#"
console.log(3.7 | 0);
console.log(-3.7 | 0);
console.log(2**32 + 1 | 0); // wraps around int32
"#), vec!["3", "-3", "1"]);
}
