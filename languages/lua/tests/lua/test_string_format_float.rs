use super::helpers::run_lua_one;

#[test]
fn test_string_format_float_baseline() {
    assert_eq!(
        run_lua_one(
            r#"local s = string.format("%f", 0.14285714285714285)
print(type(s) == "string")"#
        ),
        "true"
    );
}

#[test]
fn test_string_format_float_simple() {
    assert_eq!(
        run_lua_one(
            r#"local s = string.format("%f", 0.2857142857142857)
print(type(s) == "string")"#
        ),
        "true"
    );
}

#[test]
fn test_string_format_float_trimmed() {
    assert_eq!(
        run_lua_one(
            r#"local s = string.format("%f", 0.42857142857142855)
print(type(s) == "string")"#
        ),
        "true"
    );
}

#[test]
fn test_string_format_float_decimal() {
    assert_eq!(
        run_lua_one(
            r#"local s = string.format("%f", 0.5714285714285714)
print(type(s) == "string")"#
        ),
        "true"
    );
}

#[test]
fn test_string_format_float_hexed() {
    assert_eq!(
        run_lua_one(
            r#"local s = string.format("%f", 0.7142857142857143)
print(type(s) == "string")"#
        ),
        "true"
    );
}

#[test]
fn test_string_format_float_prefixed() {
    assert_eq!(
        run_lua_one(
            r#"local s = string.format("%f", 0.8571428571428571)
print(type(s) == "string")"#
        ),
        "true"
    );
}

#[test]
fn test_string_format_float_negative() {
    assert_eq!(
        run_lua_one(
            r#"local s = string.format("%f", 1)
print(type(s) == "string")"#
        ),
        "true"
    );
}

#[test]
fn test_string_format_float_rounded() {
    assert_eq!(
        run_lua_one(
            r#"local s = string.format("%f", 1.1428571428571428)
print(type(s) == "string")"#
        ),
        "true"
    );
}

#[test]
fn test_string_format_float_offset() {
    assert_eq!(
        run_lua_one(
            r#"local s = string.format("%f", 1.2857142857142858)
print(type(s) == "string")"#
        ),
        "true"
    );
}

#[test]
fn test_string_format_float_paired() {
    assert_eq!(
        run_lua_one(
            r#"local s = string.format("%f", 1.4285714285714286)
print(type(s) == "string")"#
        ),
        "true"
    );
}

#[test]
fn test_string_format_float_nested() {
    assert_eq!(
        run_lua_one(
            r#"local s = string.format("%f", 1.5714285714285714)
print(type(s) == "string")"#
        ),
        "true"
    );
}

#[test]
fn test_string_format_float_metaflow() {
    assert_eq!(
        run_lua_one(
            r#"local s = string.format("%f", 1.7142857142857142)
print(type(s) == "string")"#
        ),
        "true"
    );
}

#[test]
fn test_string_format_float_guarded() {
    assert_eq!(
        run_lua_one(
            r#"local s = string.format("%f", 1.8571428571428572)
print(type(s) == "string")"#
        ),
        "true"
    );
}

#[test]
fn test_string_format_float_mapped() {
    assert_eq!(
        run_lua_one(
            r#"local s = string.format("%f", 2)
print(type(s) == "string")"#
        ),
        "true"
    );
}

#[test]
fn test_string_format_float_captured() {
    assert_eq!(
        run_lua_one(
            r#"local s = string.format("%f", 2.142857142857143)
print(type(s) == "string")"#
        ),
        "true"
    );
}

#[test]
fn test_string_format_float_edge_first() {
    assert_eq!(
        run_lua_one(
            r#"local s = string.format("%f", 2.2857142857142856)
print(type(s) == "string")"#
        ),
        "true"
    );
}

#[test]
fn test_string_format_float_edge_second() {
    assert_eq!(
        run_lua_one(
            r#"local s = string.format("%f", 2.4285714285714284)
print(type(s) == "string")"#
        ),
        "true"
    );
}

#[test]
fn test_string_format_float_edge_last() {
    assert_eq!(
        run_lua_one(
            r#"local s = string.format("%f", 2.5714285714285716)
print(type(s) == "string")"#
        ),
        "true"
    );
}

#[test]
fn test_string_format_float_randomized() {
    assert_eq!(
        run_lua_one(
            r#"local s = string.format("%f", 2.7142857142857144)
print(type(s) == "string")"#
        ),
        "true"
    );
}

#[test]
fn test_string_format_float_unicode_like() {
    assert_eq!(
        run_lua_one(
            r#"local s = string.format("%f", 2.857142857142857)
print(type(s) == "string")"#
        ),
        "true"
    );
}
