use super::helpers::run_lua_one;

#[test]
fn test_print_multiple_baseline() {
    assert_eq!(run_lua_one(r#"print("a", 1, true)"#), "a 1 true");
}

#[test]
fn test_print_multiple_simple() {
    assert_eq!(run_lua_one(r#"print("b", 2, false)"#), "b 2 false");
}

#[test]
fn test_print_multiple_trimmed() {
    assert_eq!(run_lua_one(r#"print(1, 2, 3)"#), "1 2 3");
}

#[test]
fn test_print_multiple_decimal() {
    assert_eq!(run_lua_one(r#"print(nil, "x")"#), "nil x");
}

#[test]
fn test_print_multiple_hexed() {
    assert_eq!(run_lua_one(r#"print("k", nil, 4)"#), "k nil 4");
}

#[test]
fn test_print_multiple_prefixed() {
    assert_eq!(run_lua_one(r#"print("m", "n", "o")"#), "m n o");
}

#[test]
fn test_print_multiple_negative() {
    assert_eq!(run_lua_one(r#"print(5, 6, nil)"#), "5 6 nil");
}

#[test]
fn test_print_multiple_rounded() {
    assert_eq!(run_lua_one(r#"print("t", true, true)"#), "t true true");
}

#[test]
fn test_print_multiple_offset() {
    assert_eq!(run_lua_one(r#"print("r", false, 9)"#), "r false 9");
}

#[test]
fn test_print_multiple_paired() {
    assert_eq!(run_lua_one(r#"print(0, "z")"#), "0 z");
}

#[test]
fn test_print_multiple_nested() {
    assert_eq!(run_lua_one(r#"print("aa", "bb")"#), "aa bb");
}

#[test]
fn test_print_multiple_metaflow() {
    assert_eq!(run_lua_one(r#"print("u", 7, 8, 9)"#), "u 7 8 9");
}

#[test]
fn test_print_multiple_guarded() {
    assert_eq!(run_lua_one(r#"print("multi", 10, 11)"#), "multi 10 11");
}

#[test]
fn test_print_multiple_mapped() {
    assert_eq!(run_lua_one(r#"print("x", "y", nil)"#), "x y nil");
}

#[test]
fn test_print_multiple_captured() {
    assert_eq!(
        run_lua_one(r#"print("alpha", "beta", "gamma")"#),
        "alpha beta gamma"
    );
}

#[test]
fn test_print_multiple_edge_first() {
    assert_eq!(run_lua_one(r#"print("left", "right")"#), "left right");
}

#[test]
fn test_print_multiple_edge_second() {
    assert_eq!(run_lua_one(r#"print("one", 1)"#), "one 1");
}

#[test]
fn test_print_multiple_edge_last() {
    assert_eq!(run_lua_one(r#"print("two", 2)"#), "two 2");
}

#[test]
fn test_print_multiple_randomized() {
    assert_eq!(run_lua_one(r#"print("three", 3)"#), "three 3");
}

#[test]
fn test_print_multiple_unicode_like() {
    assert_eq!(run_lua_one(r#"print("four", 4)"#), "four 4");
}
