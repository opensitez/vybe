use super::helpers::run_lua_one;

#[test]
fn test_print_multiple_baseline() {
    assert_eq!(run_lua_one(r#"print("a", 1, true)"#), "a\t1\ttrue");
}

#[test]
fn test_print_multiple_simple() {
    assert_eq!(run_lua_one(r#"print("b", 2, false)"#), "b\t2\tfalse");
}

#[test]
fn test_print_multiple_trimmed() {
    assert_eq!(run_lua_one(r#"print(1, 2, 3)"#), "1\t2\t3");
}

#[test]
fn test_print_multiple_decimal() {
    assert_eq!(run_lua_one(r#"print(nil, "x")"#), "nil\tx");
}

#[test]
fn test_print_multiple_hexed() {
    assert_eq!(run_lua_one(r#"print("k", nil, 4)"#), "k\tnil\t4");
}

#[test]
fn test_print_multiple_prefixed() {
    assert_eq!(run_lua_one(r#"print("m", "n", "o")"#), "m\tn\to");
}

#[test]
fn test_print_multiple_negative() {
    assert_eq!(run_lua_one(r#"print(5, 6, nil)"#), "5\t6\tnil");
}

#[test]
fn test_print_multiple_rounded() {
    assert_eq!(run_lua_one(r#"print("t", true, true)"#), "t\ttrue\ttrue");
}

#[test]
fn test_print_multiple_offset() {
    assert_eq!(run_lua_one(r#"print("r", false, 9)"#), "r\tfalse\t9");
}

#[test]
fn test_print_multiple_paired() {
    assert_eq!(run_lua_one(r#"print(0, "z")"#), "0\tz");
}

#[test]
fn test_print_multiple_nested() {
    assert_eq!(run_lua_one(r#"print("aa", "bb")"#), "aa\tbb");
}

#[test]
fn test_print_multiple_metaflow() {
    assert_eq!(run_lua_one(r#"print("u", 7, 8, 9)"#), "u\t7\t8\t9");
}

#[test]
fn test_print_multiple_guarded() {
    assert_eq!(run_lua_one(r#"print("multi", 10, 11)"#), "multi\t10\t11");
}

#[test]
fn test_print_multiple_mapped() {
    assert_eq!(run_lua_one(r#"print("x", "y", nil)"#), "x\ty\tnil");
}

#[test]
fn test_print_multiple_captured() {
    assert_eq!(
        run_lua_one(r#"print("alpha", "beta", "gamma")"#),
        "alpha\tbeta\tgamma"
    );
}

#[test]
fn test_print_multiple_edge_first() {
    assert_eq!(run_lua_one(r#"print("left", "right")"#), "left\tright");
}

#[test]
fn test_print_multiple_edge_second() {
    assert_eq!(run_lua_one(r#"print("one", 1)"#), "one\t1");
}

#[test]
fn test_print_multiple_edge_last() {
    assert_eq!(run_lua_one(r#"print("two", 2)"#), "two\t2");
}

#[test]
fn test_print_multiple_randomized() {
    assert_eq!(run_lua_one(r#"print("three", 3)"#), "three\t3");
}

#[test]
fn test_print_multiple_unicode_like() {
    assert_eq!(run_lua_one(r#"print("four", 4)"#), "four\t4");
}
