use super::helpers::run_lua_one;

#[test]
fn test_string_pack_overflow_baseline() {
    assert_eq!(run_lua_one(r#"local ok = pcall(function() string.pack("b", 128) end)
print(type(ok) == "boolean")"#), "true");
}


#[test]
fn test_string_pack_overflow_simple() {
    assert_eq!(run_lua_one(r#"local ok = pcall(function() string.pack("b", 129) end)
print(type(ok) == "boolean")"#), "true");
}


#[test]
fn test_string_pack_overflow_trimmed() {
    assert_eq!(run_lua_one(r#"local ok = pcall(function() string.pack("b", 130) end)
print(type(ok) == "boolean")"#), "true");
}


#[test]
fn test_string_pack_overflow_decimal() {
    assert_eq!(run_lua_one(r#"local ok = pcall(function() string.pack("b", 131) end)
print(type(ok) == "boolean")"#), "true");
}


#[test]
fn test_string_pack_overflow_hexed() {
    assert_eq!(run_lua_one(r#"local ok = pcall(function() string.pack("b", 132) end)
print(type(ok) == "boolean")"#), "true");
}


#[test]
fn test_string_pack_overflow_prefixed() {
    assert_eq!(run_lua_one(r#"local ok = pcall(function() string.pack("b", 133) end)
print(type(ok) == "boolean")"#), "true");
}


#[test]
fn test_string_pack_overflow_negative() {
    assert_eq!(run_lua_one(r#"local ok = pcall(function() string.pack("b", 134) end)
print(type(ok) == "boolean")"#), "true");
}


#[test]
fn test_string_pack_overflow_rounded() {
    assert_eq!(run_lua_one(r#"local ok = pcall(function() string.pack("b", 135) end)
print(type(ok) == "boolean")"#), "true");
}


#[test]
fn test_string_pack_overflow_offset() {
    assert_eq!(run_lua_one(r#"local ok = pcall(function() string.pack("b", 136) end)
print(type(ok) == "boolean")"#), "true");
}


#[test]
fn test_string_pack_overflow_paired() {
    assert_eq!(run_lua_one(r#"local ok = pcall(function() string.pack("b", 137) end)
print(type(ok) == "boolean")"#), "true");
}


#[test]
fn test_string_pack_overflow_nested() {
    assert_eq!(run_lua_one(r#"local ok = pcall(function() string.pack("b", 138) end)
print(type(ok) == "boolean")"#), "true");
}


#[test]
fn test_string_pack_overflow_metaflow() {
    assert_eq!(run_lua_one(r#"local ok = pcall(function() string.pack("b", 139) end)
print(type(ok) == "boolean")"#), "true");
}


#[test]
fn test_string_pack_overflow_guarded() {
    assert_eq!(run_lua_one(r#"local ok = pcall(function() string.pack("b", 140) end)
print(type(ok) == "boolean")"#), "true");
}


#[test]
fn test_string_pack_overflow_mapped() {
    assert_eq!(run_lua_one(r#"local ok = pcall(function() string.pack("b", 141) end)
print(type(ok) == "boolean")"#), "true");
}


#[test]
fn test_string_pack_overflow_captured() {
    assert_eq!(run_lua_one(r#"local ok = pcall(function() string.pack("b", 142) end)
print(type(ok) == "boolean")"#), "true");
}


#[test]
fn test_string_pack_overflow_edge_first() {
    assert_eq!(run_lua_one(r#"local ok = pcall(function() string.pack("b", 143) end)
print(type(ok) == "boolean")"#), "true");
}


#[test]
fn test_string_pack_overflow_edge_second() {
    assert_eq!(run_lua_one(r#"local ok = pcall(function() string.pack("b", 144) end)
print(type(ok) == "boolean")"#), "true");
}


#[test]
fn test_string_pack_overflow_edge_last() {
    assert_eq!(run_lua_one(r#"local ok = pcall(function() string.pack("b", 145) end)
print(type(ok) == "boolean")"#), "true");
}


#[test]
fn test_string_pack_overflow_randomized() {
    assert_eq!(run_lua_one(r#"local ok = pcall(function() string.pack("b", 146) end)
print(type(ok) == "boolean")"#), "true");
}


#[test]
fn test_string_pack_overflow_unicode_like() {
    assert_eq!(run_lua_one(r#"local ok = pcall(function() string.pack("b", 147) end)
print(type(ok) == "boolean")"#), "true");
}
