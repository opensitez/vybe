use super::helpers::run_lua_one;

#[test]
fn test_string_sub_step_baseline() {
    assert_eq!(run_lua_one(r#"local t = {}
for i = 1, 6 do t[i] = string.sub("abcdefghijklmnopqrstuvwxyz", i, i) end
print(#t == 6)"#), "true");
}


#[test]
fn test_string_sub_step_simple() {
    assert_eq!(run_lua_one(r#"local t = {}
for i = 1, 7 do t[i] = string.sub("abcdefghijklmnopqrstuvwxyz", i, i) end
print(#t == 7)"#), "true");
}


#[test]
fn test_string_sub_step_trimmed() {
    assert_eq!(run_lua_one(r#"local t = {}
for i = 1, 8 do t[i] = string.sub("abcdefghijklmnopqrstuvwxyz", i, i) end
print(#t == 8)"#), "true");
}


#[test]
fn test_string_sub_step_decimal() {
    assert_eq!(run_lua_one(r#"local t = {}
for i = 1, 9 do t[i] = string.sub("abcdefghijklmnopqrstuvwxyz", i, i) end
print(#t == 9)"#), "true");
}


#[test]
fn test_string_sub_step_hexed() {
    assert_eq!(run_lua_one(r#"local t = {}
for i = 1, 10 do t[i] = string.sub("abcdefghijklmnopqrstuvwxyz", i, i) end
print(#t == 10)"#), "true");
}


#[test]
fn test_string_sub_step_prefixed() {
    assert_eq!(run_lua_one(r#"local t = {}
for i = 1, 11 do t[i] = string.sub("abcdefghijklmnopqrstuvwxyz", i, i) end
print(#t == 11)"#), "true");
}


#[test]
fn test_string_sub_step_negative() {
    assert_eq!(run_lua_one(r#"local t = {}
for i = 1, 12 do t[i] = string.sub("abcdefghijklmnopqrstuvwxyz", i, i) end
print(#t == 12)"#), "true");
}


#[test]
fn test_string_sub_step_rounded() {
    assert_eq!(run_lua_one(r#"local t = {}
for i = 1, 13 do t[i] = string.sub("abcdefghijklmnopqrstuvwxyz", i, i) end
print(#t == 13)"#), "true");
}


#[test]
fn test_string_sub_step_offset() {
    assert_eq!(run_lua_one(r#"local t = {}
for i = 1, 14 do t[i] = string.sub("abcdefghijklmnopqrstuvwxyz", i, i) end
print(#t == 14)"#), "true");
}


#[test]
fn test_string_sub_step_paired() {
    assert_eq!(run_lua_one(r#"local t = {}
for i = 1, 15 do t[i] = string.sub("abcdefghijklmnopqrstuvwxyz", i, i) end
print(#t == 15)"#), "true");
}


#[test]
fn test_string_sub_step_nested() {
    assert_eq!(run_lua_one(r#"local t = {}
for i = 1, 16 do t[i] = string.sub("abcdefghijklmnopqrstuvwxyz", i, i) end
print(#t == 16)"#), "true");
}


#[test]
fn test_string_sub_step_metaflow() {
    assert_eq!(run_lua_one(r#"local t = {}
for i = 1, 17 do t[i] = string.sub("abcdefghijklmnopqrstuvwxyz", i, i) end
print(#t == 17)"#), "true");
}


#[test]
fn test_string_sub_step_guarded() {
    assert_eq!(run_lua_one(r#"local t = {}
for i = 1, 18 do t[i] = string.sub("abcdefghijklmnopqrstuvwxyz", i, i) end
print(#t == 18)"#), "true");
}


#[test]
fn test_string_sub_step_mapped() {
    assert_eq!(run_lua_one(r#"local t = {}
for i = 1, 19 do t[i] = string.sub("abcdefghijklmnopqrstuvwxyz", i, i) end
print(#t == 19)"#), "true");
}


#[test]
fn test_string_sub_step_captured() {
    assert_eq!(run_lua_one(r#"local t = {}
for i = 1, 20 do t[i] = string.sub("abcdefghijklmnopqrstuvwxyz", i, i) end
print(#t == 20)"#), "true");
}


#[test]
fn test_string_sub_step_edge_first() {
    assert_eq!(run_lua_one(r#"local t = {}
for i = 1, 21 do t[i] = string.sub("abcdefghijklmnopqrstuvwxyz", i, i) end
print(#t == 21)"#), "true");
}


#[test]
fn test_string_sub_step_edge_second() {
    assert_eq!(run_lua_one(r#"local t = {}
for i = 1, 22 do t[i] = string.sub("abcdefghijklmnopqrstuvwxyz", i, i) end
print(#t == 22)"#), "true");
}


#[test]
fn test_string_sub_step_edge_last() {
    assert_eq!(run_lua_one(r#"local t = {}
for i = 1, 23 do t[i] = string.sub("abcdefghijklmnopqrstuvwxyz", i, i) end
print(#t == 23)"#), "true");
}


#[test]
fn test_string_sub_step_randomized() {
    assert_eq!(run_lua_one(r#"local t = {}
for i = 1, 24 do t[i] = string.sub("abcdefghijklmnopqrstuvwxyz", i, i) end
print(#t == 24)"#), "true");
}


#[test]
fn test_string_sub_step_unicode_like() {
    assert_eq!(run_lua_one(r#"local t = {}
for i = 1, 25 do t[i] = string.sub("abcdefghijklmnopqrstuvwxyz", i, i) end
print(#t == 25)"#), "true");
}
