use super::helpers::run_lua_one;

#[test]
fn test_math_modf_cases_baseline() {
    assert_eq!(
        run_lua_one(
            r#"local i, f = math.modf(1.75)
print(i == 1 and f > 0)"#
        ),
        "true"
    );
}

#[test]
fn test_math_modf_cases_simple() {
    assert_eq!(
        run_lua_one(
            r#"local i, f = math.modf(2.75)
print(i == 2 and f > 0)"#
        ),
        "true"
    );
}

#[test]
fn test_math_modf_cases_trimmed() {
    assert_eq!(
        run_lua_one(
            r#"local i, f = math.modf(3.75)
print(i == 3 and f > 0)"#
        ),
        "true"
    );
}

#[test]
fn test_math_modf_cases_decimal() {
    assert_eq!(
        run_lua_one(
            r#"local i, f = math.modf(4.75)
print(i == 4 and f > 0)"#
        ),
        "true"
    );
}

#[test]
fn test_math_modf_cases_hexed() {
    assert_eq!(
        run_lua_one(
            r#"local i, f = math.modf(5.75)
print(i == 5 and f > 0)"#
        ),
        "true"
    );
}

#[test]
fn test_math_modf_cases_prefixed() {
    assert_eq!(
        run_lua_one(
            r#"local i, f = math.modf(6.75)
print(i == 6 and f > 0)"#
        ),
        "true"
    );
}

#[test]
fn test_math_modf_cases_negative() {
    assert_eq!(
        run_lua_one(
            r#"local i, f = math.modf(7.75)
print(i == 7 and f > 0)"#
        ),
        "true"
    );
}

#[test]
fn test_math_modf_cases_rounded() {
    assert_eq!(
        run_lua_one(
            r#"local i, f = math.modf(8.75)
print(i == 8 and f > 0)"#
        ),
        "true"
    );
}

#[test]
fn test_math_modf_cases_offset() {
    assert_eq!(
        run_lua_one(
            r#"local i, f = math.modf(9.75)
print(i == 9 and f > 0)"#
        ),
        "true"
    );
}

#[test]
fn test_math_modf_cases_paired() {
    assert_eq!(
        run_lua_one(
            r#"local i, f = math.modf(10.75)
print(i == 10 and f > 0)"#
        ),
        "true"
    );
}

#[test]
fn test_math_modf_cases_nested() {
    assert_eq!(
        run_lua_one(
            r#"local i, f = math.modf(11.75)
print(i == 11 and f > 0)"#
        ),
        "true"
    );
}

#[test]
fn test_math_modf_cases_metaflow() {
    assert_eq!(
        run_lua_one(
            r#"local i, f = math.modf(12.75)
print(i == 12 and f > 0)"#
        ),
        "true"
    );
}

#[test]
fn test_math_modf_cases_guarded() {
    assert_eq!(
        run_lua_one(
            r#"local i, f = math.modf(13.75)
print(i == 13 and f > 0)"#
        ),
        "true"
    );
}

#[test]
fn test_math_modf_cases_mapped() {
    assert_eq!(
        run_lua_one(
            r#"local i, f = math.modf(14.75)
print(i == 14 and f > 0)"#
        ),
        "true"
    );
}

#[test]
fn test_math_modf_cases_captured() {
    assert_eq!(
        run_lua_one(
            r#"local i, f = math.modf(15.75)
print(i == 15 and f > 0)"#
        ),
        "true"
    );
}

#[test]
fn test_math_modf_cases_edge_first() {
    assert_eq!(
        run_lua_one(
            r#"local i, f = math.modf(16.75)
print(i == 16 and f > 0)"#
        ),
        "true"
    );
}

#[test]
fn test_math_modf_cases_edge_second() {
    assert_eq!(
        run_lua_one(
            r#"local i, f = math.modf(17.75)
print(i == 17 and f > 0)"#
        ),
        "true"
    );
}

#[test]
fn test_math_modf_cases_edge_last() {
    assert_eq!(
        run_lua_one(
            r#"local i, f = math.modf(18.75)
print(i == 18 and f > 0)"#
        ),
        "true"
    );
}

#[test]
fn test_math_modf_cases_randomized() {
    assert_eq!(
        run_lua_one(
            r#"local i, f = math.modf(19.75)
print(i == 19 and f > 0)"#
        ),
        "true"
    );
}

#[test]
fn test_math_modf_cases_unicode_like() {
    assert_eq!(
        run_lua_one(
            r#"local i, f = math.modf(20.75)
print(i == 20 and f > 0)"#
        ),
        "true"
    );
}
