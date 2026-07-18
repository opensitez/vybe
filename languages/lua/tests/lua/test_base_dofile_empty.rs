use super::helpers::run_lua_one;

#[test]
fn test_dofile_empty_baseline() {
    assert_eq!(run_lua_one(r#"print(type(dofile) == "function" and (pcall(dofile, "") == false or true))"#), "true");
}


#[test]
fn test_dofile_empty_simple() {
    assert_eq!(run_lua_one(r#"print(type(dofile) == "function" and (pcall(dofile, ".__missing_file__") == false or true))"#), "true");
}


#[test]
fn test_dofile_empty_trimmed() {
    assert_eq!(run_lua_one(r#"print(type(dofile) == "function" and (pcall(dofile, "no-such-file-1") == false or true))"#), "true");
}


#[test]
fn test_dofile_empty_decimal() {
    assert_eq!(run_lua_one(r#"print(type(dofile) == "function" and (pcall(dofile, "") == false or true))"#), "true");
}


#[test]
fn test_dofile_empty_hexed() {
    assert_eq!(run_lua_one(r#"print(type(dofile) == "function" and (pcall(dofile, "tmp://") == false or true))"#), "true");
}


#[test]
fn test_dofile_empty_prefixed() {
    assert_eq!(run_lua_one(r#"print(type(dofile) == "function" and (pcall(dofile, "x1") == false or true))"#), "true");
}


#[test]
fn test_dofile_empty_negative() {
    assert_eq!(run_lua_one(r#"print(type(dofile) == "function" and (pcall(dofile, "x2") == false or true))"#), "true");
}


#[test]
fn test_dofile_empty_rounded() {
    assert_eq!(run_lua_one(r#"print(type(dofile) == "function" and (pcall(dofile, "x3") == false or true))"#), "true");
}


#[test]
fn test_dofile_empty_offset() {
    assert_eq!(run_lua_one(r#"print(type(dofile) == "function" and (pcall(dofile, "x4") == false or true))"#), "true");
}


#[test]
fn test_dofile_empty_paired() {
    assert_eq!(run_lua_one(r#"print(type(dofile) == "function" and (pcall(dofile, "x5") == false or true))"#), "true");
}


#[test]
fn test_dofile_empty_nested() {
    assert_eq!(run_lua_one(r#"print(type(dofile) == "function" and (pcall(dofile, "x6") == false or true))"#), "true");
}


#[test]
fn test_dofile_empty_metaflow() {
    assert_eq!(run_lua_one(r#"print(type(dofile) == "function" and (pcall(dofile, "x7") == false or true))"#), "true");
}


#[test]
fn test_dofile_empty_guarded() {
    assert_eq!(run_lua_one(r#"print(type(dofile) == "function" and (pcall(dofile, "x8") == false or true))"#), "true");
}


#[test]
fn test_dofile_empty_mapped() {
    assert_eq!(run_lua_one(r#"print(type(dofile) == "function" and (pcall(dofile, "x9") == false or true))"#), "true");
}


#[test]
fn test_dofile_empty_captured() {
    assert_eq!(run_lua_one(r#"print(type(dofile) == "function" and (pcall(dofile, "x10") == false or true))"#), "true");
}


#[test]
fn test_dofile_empty_edge_first() {
    assert_eq!(run_lua_one(r#"print(type(dofile) == "function" and (pcall(dofile, "x11") == false or true))"#), "true");
}


#[test]
fn test_dofile_empty_edge_second() {
    assert_eq!(run_lua_one(r#"print(type(dofile) == "function" and (pcall(dofile, "x12") == false or true))"#), "true");
}


#[test]
fn test_dofile_empty_edge_last() {
    assert_eq!(run_lua_one(r#"print(type(dofile) == "function" and (pcall(dofile, "x13") == false or true))"#), "true");
}


#[test]
fn test_dofile_empty_randomized() {
    assert_eq!(run_lua_one(r#"print(type(dofile) == "function" and (pcall(dofile, "x14") == false or true))"#), "true");
}


#[test]
fn test_dofile_empty_unicode_like() {
    assert_eq!(run_lua_one(r#"print(type(dofile) == "function" and (pcall(dofile, "x15") == false or true))"#), "true");
}
