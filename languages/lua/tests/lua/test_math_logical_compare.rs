use super::helpers::run_lua_one;

#[test]
fn test_math_logical_compare_baseline() {
    assert_eq!(
        run_lua_one(r#"print(((0 < 1) and (1 > 0) and (0 <= 1) and (1 >= 0)))"#),
        "true"
    );
}

#[test]
fn test_math_logical_compare_simple() {
    assert_eq!(
        run_lua_one(r#"print(((1 < 2) and (2 > 1) and (1 <= 2) and (2 >= 1)))"#),
        "true"
    );
}

#[test]
fn test_math_logical_compare_trimmed() {
    assert_eq!(
        run_lua_one(r#"print(((2 < 3) and (3 > 2) and (2 <= 3) and (3 >= 2)))"#),
        "true"
    );
}

#[test]
fn test_math_logical_compare_decimal() {
    assert_eq!(
        run_lua_one(r#"print(((3 < 4) and (4 > 3) and (3 <= 4) and (4 >= 3)))"#),
        "true"
    );
}

#[test]
fn test_math_logical_compare_hexed() {
    assert_eq!(
        run_lua_one(r#"print(((4 < 5) and (5 > 4) and (4 <= 5) and (5 >= 4)))"#),
        "true"
    );
}

#[test]
fn test_math_logical_compare_prefixed() {
    assert_eq!(
        run_lua_one(r#"print(((5 < 6) and (6 > 5) and (5 <= 6) and (6 >= 5)))"#),
        "true"
    );
}

#[test]
fn test_math_logical_compare_negative() {
    assert_eq!(
        run_lua_one(r#"print(((6 < 7) and (7 > 6) and (6 <= 7) and (7 >= 6)))"#),
        "true"
    );
}

#[test]
fn test_math_logical_compare_rounded() {
    assert_eq!(
        run_lua_one(r#"print(((7 < 8) and (8 > 7) and (7 <= 8) and (8 >= 7)))"#),
        "true"
    );
}

#[test]
fn test_math_logical_compare_offset() {
    assert_eq!(
        run_lua_one(r#"print(((8 < 9) and (9 > 8) and (8 <= 9) and (9 >= 8)))"#),
        "true"
    );
}

#[test]
fn test_math_logical_compare_paired() {
    assert_eq!(
        run_lua_one(r#"print(((9 < 10) and (10 > 9) and (9 <= 10) and (10 >= 9)))"#),
        "true"
    );
}

#[test]
fn test_math_logical_compare_nested() {
    assert_eq!(
        run_lua_one(r#"print(((10 < 11) and (11 > 10) and (10 <= 11) and (11 >= 10)))"#),
        "true"
    );
}

#[test]
fn test_math_logical_compare_metaflow() {
    assert_eq!(
        run_lua_one(r#"print(((11 < 12) and (12 > 11) and (11 <= 12) and (12 >= 11)))"#),
        "true"
    );
}

#[test]
fn test_math_logical_compare_guarded() {
    assert_eq!(
        run_lua_one(r#"print(((12 < 13) and (13 > 12) and (12 <= 13) and (13 >= 12)))"#),
        "true"
    );
}

#[test]
fn test_math_logical_compare_mapped() {
    assert_eq!(
        run_lua_one(r#"print(((13 < 14) and (14 > 13) and (13 <= 14) and (14 >= 13)))"#),
        "true"
    );
}

#[test]
fn test_math_logical_compare_captured() {
    assert_eq!(
        run_lua_one(r#"print(((14 < 15) and (15 > 14) and (14 <= 15) and (15 >= 14)))"#),
        "true"
    );
}

#[test]
fn test_math_logical_compare_edge_first() {
    assert_eq!(
        run_lua_one(r#"print(((15 < 16) and (16 > 15) and (15 <= 16) and (16 >= 15)))"#),
        "true"
    );
}

#[test]
fn test_math_logical_compare_edge_second() {
    assert_eq!(
        run_lua_one(r#"print(((16 < 17) and (17 > 16) and (16 <= 17) and (17 >= 16)))"#),
        "true"
    );
}

#[test]
fn test_math_logical_compare_edge_last() {
    assert_eq!(
        run_lua_one(r#"print(((17 < 18) and (18 > 17) and (17 <= 18) and (18 >= 17)))"#),
        "true"
    );
}

#[test]
fn test_math_logical_compare_randomized() {
    assert_eq!(
        run_lua_one(r#"print(((18 < 19) and (19 > 18) and (18 <= 19) and (19 >= 18)))"#),
        "true"
    );
}

#[test]
fn test_math_logical_compare_unicode_like() {
    assert_eq!(
        run_lua_one(r#"print(((19 < 20) and (20 > 19) and (19 <= 20) and (20 >= 19)))"#),
        "true"
    );
}
