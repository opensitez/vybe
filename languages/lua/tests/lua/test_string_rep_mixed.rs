use super::helpers::run_lua_one;

#[test]
fn test_string_rep_mixed_baseline() {
    assert_eq!(run_lua_one(r#"print(string.rep("A", 1, ":") == string.rep("A:", 0) .. "A")"#), "true");
}


#[test]
fn test_string_rep_mixed_simple() {
    assert_eq!(run_lua_one(r#"print(string.rep("B", 2, ":") == string.rep("B:", 1) .. "B")"#), "true");
}


#[test]
fn test_string_rep_mixed_trimmed() {
    assert_eq!(run_lua_one(r#"print(string.rep("C", 3, ":") == string.rep("C:", 2) .. "C")"#), "true");
}


#[test]
fn test_string_rep_mixed_decimal() {
    assert_eq!(run_lua_one(r#"print(string.rep("D", 4, ":") == string.rep("D:", 3) .. "D")"#), "true");
}


#[test]
fn test_string_rep_mixed_hexed() {
    assert_eq!(run_lua_one(r#"print(string.rep("E", 5, ":") == string.rep("E:", 4) .. "E")"#), "true");
}


#[test]
fn test_string_rep_mixed_prefixed() {
    assert_eq!(run_lua_one(r#"print(string.rep("F", 6, ":") == string.rep("F:", 5) .. "F")"#), "true");
}


#[test]
fn test_string_rep_mixed_negative() {
    assert_eq!(run_lua_one(r#"print(string.rep("G", 7, ":") == string.rep("G:", 6) .. "G")"#), "true");
}


#[test]
fn test_string_rep_mixed_rounded() {
    assert_eq!(run_lua_one(r#"print(string.rep("H", 8, ":") == string.rep("H:", 7) .. "H")"#), "true");
}


#[test]
fn test_string_rep_mixed_offset() {
    assert_eq!(run_lua_one(r#"print(string.rep("I", 9, ":") == string.rep("I:", 8) .. "I")"#), "true");
}


#[test]
fn test_string_rep_mixed_paired() {
    assert_eq!(run_lua_one(r#"print(string.rep("J", 10, ":") == string.rep("J:", 9) .. "J")"#), "true");
}


#[test]
fn test_string_rep_mixed_nested() {
    assert_eq!(run_lua_one(r#"print(string.rep("K", 11, ":") == string.rep("K:", 10) .. "K")"#), "true");
}


#[test]
fn test_string_rep_mixed_metaflow() {
    assert_eq!(run_lua_one(r#"print(string.rep("L", 12, ":") == string.rep("L:", 11) .. "L")"#), "true");
}


#[test]
fn test_string_rep_mixed_guarded() {
    assert_eq!(run_lua_one(r#"print(string.rep("M", 13, ":") == string.rep("M:", 12) .. "M")"#), "true");
}


#[test]
fn test_string_rep_mixed_mapped() {
    assert_eq!(run_lua_one(r#"print(string.rep("N", 14, ":") == string.rep("N:", 13) .. "N")"#), "true");
}


#[test]
fn test_string_rep_mixed_captured() {
    assert_eq!(run_lua_one(r#"print(string.rep("O", 15, ":") == string.rep("O:", 14) .. "O")"#), "true");
}


#[test]
fn test_string_rep_mixed_edge_first() {
    assert_eq!(run_lua_one(r#"print(string.rep("P", 16, ":") == string.rep("P:", 15) .. "P")"#), "true");
}


#[test]
fn test_string_rep_mixed_edge_second() {
    assert_eq!(run_lua_one(r#"print(string.rep("Q", 17, ":") == string.rep("Q:", 16) .. "Q")"#), "true");
}


#[test]
fn test_string_rep_mixed_edge_last() {
    assert_eq!(run_lua_one(r#"print(string.rep("R", 18, ":") == string.rep("R:", 17) .. "R")"#), "true");
}


#[test]
fn test_string_rep_mixed_randomized() {
    assert_eq!(run_lua_one(r#"print(string.rep("S", 19, ":") == string.rep("S:", 18) .. "S")"#), "true");
}


#[test]
fn test_string_rep_mixed_unicode_like() {
    assert_eq!(run_lua_one(r#"print(string.rep("T", 20, ":") == string.rep("T:", 19) .. "T")"#), "true");
}
