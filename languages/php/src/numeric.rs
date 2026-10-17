//! PHP numeric strings consume the entire string, unlike parseFloat.
use vybe_runtime::value::ObjectKind;
use vybe_runtime::{Framework, Value};

fn php_truthy(value: &Value) -> bool {
    match value {
        Value::Null | Value::Undefined => false,
        Value::Bool(value) => *value,
        Value::I32(value) => *value != 0,
        Value::I64(value) => *value != 0,
        Value::F32(value) => *value != 0.0 && !value.is_nan(),
        Value::F64(value) => *value != 0.0 && !value.is_nan(),
        Value::String(value) => !value.is_empty() && value.as_ref() != "0",
        Value::Object(value) => match &value.lock().unwrap().kind {
            ObjectKind::Array(items) => !items.is_empty(),
            ObjectKind::Map(items) => !items.is_empty(),
            _ => true,
        },
        _ => true,
    }
}

pub(crate) fn is_numeric_string(value: &str) -> bool {
    let value = value.trim_matches(|c| matches!(c, ' ' | '\t' | '\n' | '\r' | '\x0b' | '\x0c'));
    let bytes = value.as_bytes();
    let mut i = usize::from(matches!(bytes.first(), Some(b'+' | b'-')));
    let start = i;
    while bytes.get(i).is_some_and(u8::is_ascii_digit) {
        i += 1;
    }
    let mut digits = i - start;
    if bytes.get(i) == Some(&b'.') {
        i += 1;
        let start = i;
        while bytes.get(i).is_some_and(u8::is_ascii_digit) {
            i += 1;
        }
        digits += i - start;
    }
    if digits == 0 {
        return false;
    }
    if matches!(bytes.get(i), Some(b'e' | b'E')) {
        i += 1;
        if matches!(bytes.get(i), Some(b'+' | b'-')) {
            i += 1;
        }
        let start = i;
        while bytes.get(i).is_some_and(u8::is_ascii_digit) {
            i += 1;
        }
        if i == start {
            return false;
        }
    }
    i == bytes.len()
}

/// PHP's constant string comparison retains integer precision and falls back
/// to byte equality when integer overflow or equal infinities lose precision.
/// This is a frontend fold, not a new runtime import or VM operation.
pub(crate) fn numeric_strings_equal(left: &str, right: &str) -> bool {
    if left == right {
        return true;
    }
    if !is_numeric_string(left) || !is_numeric_string(right) {
        return false;
    }
    enum Number {
        Integer(i64),
        Float(f64, i8),
    }
    fn classify(value: &str) -> Number {
        let value = value.trim_matches(|c| matches!(c, ' ' | '\t' | '\n' | '\r' | '\x0b' | '\x0c'));
        let integer = !value.bytes().any(|byte| matches!(byte, b'.' | b'e' | b'E'));
        if integer {
            if let Ok(value) = value.parse::<i64>() {
                return Number::Integer(value);
            }
        }
        let overflow = if !integer { 0 } else if value.starts_with('-') { -1 } else { 1 };
        Number::Float(value.parse::<f64>().expect("validated decimal numeric string"), overflow)
    }
    match (classify(left), classify(right)) {
        (Number::Integer(a), Number::Integer(b)) => a == b,
        (Number::Integer(a), Number::Float(b, overflow))
        | (Number::Float(b, overflow), Number::Integer(a)) => {
            overflow == 0 && a as f64 == b
        }
        (Number::Float(a, a_overflow), Number::Float(b, b_overflow)) => {
            if a_overflow != 0 && a_overflow == b_overflow && a - b == 0.0 {
                return false;
            }
            a.is_finite() && b.is_finite() && a == b
        }
    }
}

#[cfg(test)]
mod numeric_string_equality_tests {
    use super::numeric_strings_equal;

    #[test]
    fn decimal_grammar_integer_precision_and_overflow_match_php() {
        for (left, right, expected) in [
            ("10", "10.0", true), ("\t10\n", "+10", true),
            (".5", "0.5", true), ("1e2", "100", true),
            ("1e-999", "0", true), ("-0", "+0.0", true),
            ("9a", "9", false), ("Infinity", "1e999", false),
            ("0x10", "16", false), ("1e+", "1", false),
            ("1e999", "2e999", false), ("1e999", "1e999", true),
            ("-1e999", "-2e999", false), ("1e999", "10e998", false),
            ("9223372036854775807", "9223372036854775806", false),
            ("9007199254740993", "9007199254740992", false),
            ("9223372036854775808", "9223372036854775807", false),
            ("9223372036854775808", "9223372036854775809", false),
            ("9223372036854775808", "9.223372036854776e18", true),
            ("\u{a0}10", "10", false),
        ] {
            assert_eq!(numeric_strings_equal(left, right), expected, "{left:?} == {right:?}");
            assert_eq!(numeric_strings_equal(right, left), expected, "{right:?} == {left:?}");
        }
    }
}

/// PHP's significant-digit conversion, including its exponent threshold and
/// mandatory fractional digit in scientific notation. This belongs to PHP,
/// not the VM's generic number-to-string transport.
pub(crate) fn php_float_string(value: f64, precision: i32) -> String {
    if !value.is_finite() {
        let text = if value.is_nan() { "NAN" } else if value.is_sign_negative() { "-INF" } else { "INF" };
        return if precision < 0 { text.into() } else { text.chars().take(precision.max(1) as usize).collect() };
    }
    if value == 0.0 {
        return if value.is_sign_negative() { "-0" } else { "0" }.into();
    }
    // A binary64 decimal expansion has no significant digits beyond this
    // bound. Extra requested zeros would be stripped below anyway.
    let raw = if precision < 0 {
        // Choose the shortest round-tripping precision with decimal ties
        // rounded to even. Rust's default shortest display can choose the
        // other equally short spelling at an exact decimal midpoint.
        (1..=17)
            .map(|digits| format!("{:.*e}", digits - 1, value.abs()))
            .find(|text| text.parse::<f64>().ok() == Some(value.abs()))
            .expect("binary64 decimal round trip")
    } else {
        format!("{:.*e}", precision.clamp(1, 1074) as usize - 1, value.abs())
    };
    let (mantissa, exponent) = raw.split_once('e').expect("scientific exponent");
    let exponent: i32 = exponent.parse().expect("decimal exponent");
    let mut digits = mantissa.replace('.', "");
    while digits.len() > 1 && digits.ends_with('0') { digits.pop(); }
    let limit = if precision < 0 { 17 } else { precision.max(1) };
    let mut output = if exponent < -4 || exponent >= limit {
        let fraction = if digits.len() > 1 { &digits[1..] } else { "0" };
        format!("{}.{}E{}{exponent}", &digits[..1], fraction, if exponent >= 0 { "+" } else { "" })
    } else {
        let point = exponent + 1;
        if point <= 0 {
            format!("0.{}{}", "0".repeat((-point) as usize), digits)
        } else if point as usize >= digits.len() {
            format!("{}{}", digits, "0".repeat(point as usize - digits.len()))
        } else {
            format!("{}.{}", &digits[..point as usize], &digits[point as usize..])
        }
    };
    if value.is_sign_negative() { output.insert(0, '-'); }
    output
}

pub fn register(fw: &mut Framework<'_>) {
    fw.register_host_fn(
        "php:value",
        "floatToString",
        Box::new(|_, args| {
            let number = match args.first() {
                Some(Value::F64(value)) => *value,
                Some(Value::I32(value)) => *value as f64,
                Some(Value::I64(value)) => *value as f64,
                _ => return Value::Undefined,
            };
            let precision = match args.get(1) {
                Some(Value::String(value)) => value.parse::<i32>().unwrap_or(14),
                Some(Value::I32(value)) => *value,
                Some(Value::F64(value)) => *value as i32,
                _ => 14,
            };
            Value::String(php_float_string(number, precision).into())
        }),
    );
    fw.register_host_fn(
        "php:value",
        "looseBoolEq",
        Box::new(|_, args| {
            Value::Bool(args.len() >= 2 && php_truthy(&args[0]) == php_truthy(&args[1]))
        }),
    );
    fw.register_host_fn(
        "php:value",
        "isNumeric",
        Box::new(|_, args| {
            Value::Bool(match args.first() {
                Some(Value::I32(_) | Value::I64(_) | Value::F64(_)) => true,
                Some(Value::String(value)) => is_numeric_string(value),
                _ => false,
            })
        }),
    );
}

#[cfg(test)]
mod float_string_tests {
    use super::php_float_string;
    #[test]
    fn php_exponent_precision_and_special_values() {
        for (value, precision, expected) in [
            (1e20, 14, "1.0E+20"), (1e-5, 14, "1.0E-5"),
            (1e-4, 14, "0.0001"), (1.234567890123456, 6, "1.23457"),
            (1.234567890123456, -1, "1.234567890123456"),
            (-0.0, 14, "-0"), (f64::NAN, 14, "NAN"),
            (f64::INFINITY, 14, "INF"), (f64::NEG_INFINITY, 14, "-INF"),
            (5e-324, -1, "5.0E-324"), (f64::NEG_INFINITY, 2, "-I"),
            (f64::from_bits(0x42b11802913849d0), -1, "18794819958857.812"),
        ] { assert_eq!(php_float_string(value, precision), expected); }
    }
}
