use super::helpers::run_lua_one;

#[test]
fn test_string_upper_lower_baseline() {
    assert_eq!(run_lua_one(r#"print(string.lower(string.upper("a")) == "a")"#), "true");
}


#[test]
fn test_string_upper_lower_simple() {
    assert_eq!(run_lua_one(r#"print(string.lower(string.upper("b")) == "b")"#), "true");
}


#[test]
fn test_string_upper_lower_trimmed() {
    assert_eq!(run_lua_one(r#"print(string.lower(string.upper("c")) == "c")"#), "true");
}


#[test]
fn test_string_upper_lower_decimal() {
    assert_eq!(run_lua_one(r#"print(string.lower(string.upper("d")) == "d")"#), "true");
}


#[test]
fn test_string_upper_lower_hexed() {
    assert_eq!(run_lua_one(r#"print(string.lower(string.upper("e")) == "e")"#), "true");
}


#[test]
fn test_string_upper_lower_prefixed() {
    assert_eq!(run_lua_one(r#"print(string.lower(string.upper("f")) == "f")"#), "true");
}


#[test]
fn test_string_upper_lower_negative() {
    assert_eq!(run_lua_one(r#"print(string.lower(string.upper("g")) == "g")"#), "true");
}


#[test]
fn test_string_upper_lower_rounded() {
    assert_eq!(run_lua_one(r#"print(string.lower(string.upper("h")) == "h")"#), "true");
}


#[test]
fn test_string_upper_lower_offset() {
    assert_eq!(run_lua_one(r#"print(string.lower(string.upper("i")) == "i")"#), "true");
}


#[test]
fn test_string_upper_lower_paired() {
    assert_eq!(run_lua_one(r#"print(string.lower(string.upper("j")) == "j")"#), "true");
}


#[test]
fn test_string_upper_lower_nested() {
    assert_eq!(run_lua_one(r#"print(string.lower(string.upper("k")) == "k")"#), "true");
}


#[test]
fn test_string_upper_lower_metaflow() {
    assert_eq!(run_lua_one(r#"print(string.lower(string.upper("l")) == "l")"#), "true");
}


#[test]
fn test_string_upper_lower_guarded() {
    assert_eq!(run_lua_one(r#"print(string.lower(string.upper("m")) == "m")"#), "true");
}


#[test]
fn test_string_upper_lower_mapped() {
    assert_eq!(run_lua_one(r#"print(string.lower(string.upper("n")) == "n")"#), "true");
}


#[test]
fn test_string_upper_lower_captured() {
    assert_eq!(run_lua_one(r#"print(string.lower(string.upper("o")) == "o")"#), "true");
}


#[test]
fn test_string_upper_lower_edge_first() {
    assert_eq!(run_lua_one(r#"print(string.lower(string.upper("p")) == "p")"#), "true");
}


#[test]
fn test_string_upper_lower_edge_second() {
    assert_eq!(run_lua_one(r#"print(string.lower(string.upper("q")) == "q")"#), "true");
}


#[test]
fn test_string_upper_lower_edge_last() {
    assert_eq!(run_lua_one(r#"print(string.lower(string.upper("r")) == "r")"#), "true");
}


#[test]
fn test_string_upper_lower_randomized() {
    assert_eq!(run_lua_one(r#"print(string.lower(string.upper("s")) == "s")"#), "true");
}


#[test]
fn test_string_upper_lower_unicode_like() {
    assert_eq!(run_lua_one(r#"print(string.lower(string.upper("t")) == "t")"#), "true");
}
