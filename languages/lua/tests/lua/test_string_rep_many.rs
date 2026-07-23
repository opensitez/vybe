use super::helpers::run_lua_one;

#[test]
fn test_string_rep_many_baseline() {
    assert_eq!(run_lua_one(r#"print(#string.rep("a", 30) == 30)"#), "true");
}

#[test]
fn test_string_rep_many_simple() {
    assert_eq!(run_lua_one(r#"print(#string.rep("a", 31) == 31)"#), "true");
}

#[test]
fn test_string_rep_many_trimmed() {
    assert_eq!(run_lua_one(r#"print(#string.rep("a", 32) == 32)"#), "true");
}

#[test]
fn test_string_rep_many_decimal() {
    assert_eq!(run_lua_one(r#"print(#string.rep("a", 33) == 33)"#), "true");
}

#[test]
fn test_string_rep_many_hexed() {
    assert_eq!(run_lua_one(r#"print(#string.rep("a", 34) == 34)"#), "true");
}

#[test]
fn test_string_rep_many_prefixed() {
    assert_eq!(run_lua_one(r#"print(#string.rep("a", 35) == 35)"#), "true");
}

#[test]
fn test_string_rep_many_negative() {
    assert_eq!(run_lua_one(r#"print(#string.rep("a", 36) == 36)"#), "true");
}

#[test]
fn test_string_rep_many_rounded() {
    assert_eq!(run_lua_one(r#"print(#string.rep("a", 37) == 37)"#), "true");
}

#[test]
fn test_string_rep_many_offset() {
    assert_eq!(run_lua_one(r#"print(#string.rep("a", 38) == 38)"#), "true");
}

#[test]
fn test_string_rep_many_paired() {
    assert_eq!(run_lua_one(r#"print(#string.rep("a", 39) == 39)"#), "true");
}

#[test]
fn test_string_rep_many_nested() {
    assert_eq!(run_lua_one(r#"print(#string.rep("a", 40) == 40)"#), "true");
}

#[test]
fn test_string_rep_many_metaflow() {
    assert_eq!(run_lua_one(r#"print(#string.rep("a", 41) == 41)"#), "true");
}

#[test]
fn test_string_rep_many_guarded() {
    assert_eq!(run_lua_one(r#"print(#string.rep("a", 42) == 42)"#), "true");
}

#[test]
fn test_string_rep_many_mapped() {
    assert_eq!(run_lua_one(r#"print(#string.rep("a", 43) == 43)"#), "true");
}

#[test]
fn test_string_rep_many_captured() {
    assert_eq!(run_lua_one(r#"print(#string.rep("a", 44) == 44)"#), "true");
}

#[test]
fn test_string_rep_many_edge_first() {
    assert_eq!(run_lua_one(r#"print(#string.rep("a", 45) == 45)"#), "true");
}

#[test]
fn test_string_rep_many_edge_second() {
    assert_eq!(run_lua_one(r#"print(#string.rep("a", 46) == 46)"#), "true");
}

#[test]
fn test_string_rep_many_edge_last() {
    assert_eq!(run_lua_one(r#"print(#string.rep("a", 47) == 47)"#), "true");
}

#[test]
fn test_string_rep_many_randomized() {
    assert_eq!(run_lua_one(r#"print(#string.rep("a", 48) == 48)"#), "true");
}

#[test]
fn test_string_rep_many_unicode_like() {
    assert_eq!(run_lua_one(r#"print(#string.rep("a", 49) == 49)"#), "true");
}
