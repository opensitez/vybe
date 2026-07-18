use super::helpers::run_lua_one;

#[test]
fn test_tonumber_radix_baseline() {
    assert_eq!(run_lua_one(r#"print(tonumber("1111", 2) == 15)"#), "true");
}


#[test]
fn test_tonumber_radix_simple() {
    assert_eq!(run_lua_one(r#"print(tonumber("1010", 2) == 10)"#), "true");
}


#[test]
fn test_tonumber_radix_trimmed() {
    assert_eq!(run_lua_one(r#"print(tonumber("77", 8) == 63)"#), "true");
}


#[test]
fn test_tonumber_radix_decimal() {
    assert_eq!(run_lua_one(r#"print(tonumber("10", 8) == 8)"#), "true");
}


#[test]
fn test_tonumber_radix_hexed() {
    assert_eq!(run_lua_one(r#"print(tonumber("ff", 16) == 255)"#), "true");
}


#[test]
fn test_tonumber_radix_prefixed() {
    assert_eq!(run_lua_one(r#"print(tonumber("10", 16) == 16)"#), "true");
}


#[test]
fn test_tonumber_radix_negative() {
    assert_eq!(run_lua_one(r#"print(tonumber("7f", 16) == 127)"#), "true");
}


#[test]
fn test_tonumber_radix_rounded() {
    assert_eq!(run_lua_one(r#"print(tonumber("101", 3) == 10)"#), "true");
}


#[test]
fn test_tonumber_radix_offset() {
    assert_eq!(run_lua_one(r#"print(tonumber("202", 3) == 20)"#), "true");
}


#[test]
fn test_tonumber_radix_paired() {
    assert_eq!(run_lua_one(r#"print(tonumber("20", 10) == 20)"#), "true");
}


#[test]
fn test_tonumber_radix_nested() {
    assert_eq!(run_lua_one(r#"print(tonumber("42", 10) == 42)"#), "true");
}


#[test]
fn test_tonumber_radix_metaflow() {
    assert_eq!(run_lua_one(r#"print(tonumber("-101", 2) == -5)"#), "true");
}


#[test]
fn test_tonumber_radix_guarded() {
    assert_eq!(run_lua_one(r#"print(tonumber("-ff", 16) == -255)"#), "true");
}


#[test]
fn test_tonumber_radix_mapped() {
    assert_eq!(run_lua_one(r#"print(tonumber("-10", 8) == -8)"#), "true");
}


#[test]
fn test_tonumber_radix_captured() {
    assert_eq!(run_lua_one(r#"print(tonumber("13", 16) == 19)"#), "true");
}


#[test]
fn test_tonumber_radix_edge_first() {
    assert_eq!(run_lua_one(r#"print(tonumber("12", 5) == 7)"#), "true");
}


#[test]
fn test_tonumber_radix_edge_second() {
    assert_eq!(run_lua_one(r#"print(tonumber("2a", 11) == 32)"#), "true");
}


#[test]
fn test_tonumber_radix_edge_last() {
    assert_eq!(run_lua_one(r#"print(tonumber("100", 4) == 16)"#), "true");
}


#[test]
fn test_tonumber_radix_randomized() {
    assert_eq!(run_lua_one(r#"print(tonumber("17", 7) == 1)"#), "true");
}


#[test]
fn test_tonumber_radix_unicode_like() {
    assert_eq!(run_lua_one(r#"print(tonumber("1f", 16) == 31)"#), "true");
}
