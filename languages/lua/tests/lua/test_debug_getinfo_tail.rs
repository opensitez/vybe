use super::helpers::run_lua_one;

#[test]
fn test_debug_getinfo_tail_baseline() {
    assert_eq!(run_lua_one(r#"local info = debug.getinfo(1, "n")
print(type(info) == "table")"#), "true");
}


#[test]
fn test_debug_getinfo_tail_simple() {
    assert_eq!(run_lua_one(r#"local info = debug.getinfo(1, "S")
print(type(info) == "table")"#), "true");
}


#[test]
fn test_debug_getinfo_tail_trimmed() {
    assert_eq!(run_lua_one(r#"local info = debug.getinfo(1, "f")
print(type(info) == "table")"#), "true");
}


#[test]
fn test_debug_getinfo_tail_decimal() {
    assert_eq!(run_lua_one(r#"local info = debug.getinfo(1, "l")
print(type(info) == "table")"#), "true");
}


#[test]
fn test_debug_getinfo_tail_hexed() {
    assert_eq!(run_lua_one(r#"local info = debug.getinfo(1, "u")
print(type(info) == "table")"#), "true");
}


#[test]
fn test_debug_getinfo_tail_prefixed() {
    assert_eq!(run_lua_one(r#"local info = debug.getinfo(1, "L")
print(type(info) == "table")"#), "true");
}


#[test]
fn test_debug_getinfo_tail_negative() {
    assert_eq!(run_lua_one(r#"local info = debug.getinfo(1, "nS")
print(type(info) == "table")"#), "true");
}


#[test]
fn test_debug_getinfo_tail_rounded() {
    assert_eq!(run_lua_one(r#"local info = debug.getinfo(1, "nSl")
print(type(info) == "table")"#), "true");
}


#[test]
fn test_debug_getinfo_tail_offset() {
    assert_eq!(run_lua_one(r#"local info = debug.getinfo(1, "nu")
print(type(info) == "table")"#), "true");
}


#[test]
fn test_debug_getinfo_tail_paired() {
    assert_eq!(run_lua_one(r#"local info = debug.getinfo(1, "nl")
print(type(info) == "table")"#), "true");
}


#[test]
fn test_debug_getinfo_tail_nested() {
    assert_eq!(run_lua_one(r#"local info = debug.getinfo(1, "fS")
print(type(info) == "table")"#), "true");
}


#[test]
fn test_debug_getinfo_tail_metaflow() {
    assert_eq!(run_lua_one(r#"local info = debug.getinfo(1, "fL")
print(type(info) == "table")"#), "true");
}


#[test]
fn test_debug_getinfo_tail_guarded() {
    assert_eq!(run_lua_one(r#"local info = debug.getinfo(1, "fSlu")
print(type(info) == "table")"#), "true");
}


#[test]
fn test_debug_getinfo_tail_mapped() {
    assert_eq!(run_lua_one(r#"local info = debug.getinfo(1, "nSf")
print(type(info) == "table")"#), "true");
}


#[test]
fn test_debug_getinfo_tail_captured() {
    assert_eq!(run_lua_one(r#"local info = debug.getinfo(1, "nLu")
print(type(info) == "table")"#), "true");
}


#[test]
fn test_debug_getinfo_tail_edge_first() {
    assert_eq!(run_lua_one(r#"local info = debug.getinfo(1, "nlf")
print(type(info) == "table")"#), "true");
}


#[test]
fn test_debug_getinfo_tail_edge_second() {
    assert_eq!(run_lua_one(r#"local info = debug.getinfo(1, "Slu")
print(type(info) == "table")"#), "true");
}


#[test]
fn test_debug_getinfo_tail_edge_last() {
    assert_eq!(run_lua_one(r#"local info = debug.getinfo(1, "nSu")
print(type(info) == "table")"#), "true");
}


#[test]
fn test_debug_getinfo_tail_randomized() {
    assert_eq!(run_lua_one(r#"local info = debug.getinfo(1, "nSL")
print(type(info) == "table")"#), "true");
}


#[test]
fn test_debug_getinfo_tail_unicode_like() {
    assert_eq!(run_lua_one(r#"local info = debug.getinfo(1, "nSU")
print(type(info) == "table")"#), "true");
}
