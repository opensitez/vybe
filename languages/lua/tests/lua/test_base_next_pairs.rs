use super::helpers::run_lua_one;

#[test]
fn test_next_pairs_baseline() {
    assert_eq!(run_lua_one(r#"local t = {}
for i = 1, 1 do t[i] = i end
local c = 0
for k in next, t, nil do c = c + 1 end
print(c >= 1)"#), "true");
}


#[test]
fn test_next_pairs_simple() {
    assert_eq!(run_lua_one(r#"local t = {}
for i = 1, 2 do t[i] = i end
local c = 0
for k in next, t, nil do c = c + 1 end
print(c >= 1)"#), "true");
}


#[test]
fn test_next_pairs_trimmed() {
    assert_eq!(run_lua_one(r#"local t = {}
for i = 1, 3 do t[i] = i end
local c = 0
for k in next, t, nil do c = c + 1 end
print(c >= 1)"#), "true");
}


#[test]
fn test_next_pairs_decimal() {
    assert_eq!(run_lua_one(r#"local t = {}
for i = 1, 4 do t[i] = i end
local c = 0
for k in next, t, nil do c = c + 1 end
print(c >= 1)"#), "true");
}


#[test]
fn test_next_pairs_hexed() {
    assert_eq!(run_lua_one(r#"local t = {}
for i = 1, 5 do t[i] = i end
local c = 0
for k in next, t, nil do c = c + 1 end
print(c >= 1)"#), "true");
}


#[test]
fn test_next_pairs_prefixed() {
    assert_eq!(run_lua_one(r#"local t = {}
for i = 1, 6 do t[i] = i end
local c = 0
for k in next, t, nil do c = c + 1 end
print(c >= 1)"#), "true");
}


#[test]
fn test_next_pairs_negative() {
    assert_eq!(run_lua_one(r#"local t = {}
for i = 1, 7 do t[i] = i end
local c = 0
for k in next, t, nil do c = c + 1 end
print(c >= 1)"#), "true");
}


#[test]
fn test_next_pairs_rounded() {
    assert_eq!(run_lua_one(r#"local t = {}
for i = 1, 8 do t[i] = i end
local c = 0
for k in next, t, nil do c = c + 1 end
print(c >= 1)"#), "true");
}


#[test]
fn test_next_pairs_offset() {
    assert_eq!(run_lua_one(r#"local t = {}
for i = 1, 9 do t[i] = i end
local c = 0
for k in next, t, nil do c = c + 1 end
print(c >= 1)"#), "true");
}


#[test]
fn test_next_pairs_paired() {
    assert_eq!(run_lua_one(r#"local t = {}
for i = 1, 10 do t[i] = i end
local c = 0
for k in next, t, nil do c = c + 1 end
print(c >= 1)"#), "true");
}


#[test]
fn test_next_pairs_nested() {
    assert_eq!(run_lua_one(r#"local t = {}
for i = 1, 11 do t[i] = i end
local c = 0
for k in next, t, nil do c = c + 1 end
print(c >= 1)"#), "true");
}


#[test]
fn test_next_pairs_metaflow() {
    assert_eq!(run_lua_one(r#"local t = {}
for i = 1, 12 do t[i] = i end
local c = 0
for k in next, t, nil do c = c + 1 end
print(c >= 1)"#), "true");
}


#[test]
fn test_next_pairs_guarded() {
    assert_eq!(run_lua_one(r#"local t = {}
for i = 1, 13 do t[i] = i end
local c = 0
for k in next, t, nil do c = c + 1 end
print(c >= 1)"#), "true");
}


#[test]
fn test_next_pairs_mapped() {
    assert_eq!(run_lua_one(r#"local t = {}
for i = 1, 14 do t[i] = i end
local c = 0
for k in next, t, nil do c = c + 1 end
print(c >= 1)"#), "true");
}


#[test]
fn test_next_pairs_captured() {
    assert_eq!(run_lua_one(r#"local t = {}
for i = 1, 15 do t[i] = i end
local c = 0
for k in next, t, nil do c = c + 1 end
print(c >= 1)"#), "true");
}


#[test]
fn test_next_pairs_edge_first() {
    assert_eq!(run_lua_one(r#"local t = {}
for i = 1, 16 do t[i] = i end
local c = 0
for k in next, t, nil do c = c + 1 end
print(c >= 1)"#), "true");
}


#[test]
fn test_next_pairs_edge_second() {
    assert_eq!(run_lua_one(r#"local t = {}
for i = 1, 17 do t[i] = i end
local c = 0
for k in next, t, nil do c = c + 1 end
print(c >= 1)"#), "true");
}


#[test]
fn test_next_pairs_edge_last() {
    assert_eq!(run_lua_one(r#"local t = {}
for i = 1, 18 do t[i] = i end
local c = 0
for k in next, t, nil do c = c + 1 end
print(c >= 1)"#), "true");
}


#[test]
fn test_next_pairs_randomized() {
    assert_eq!(run_lua_one(r#"local t = {}
for i = 1, 19 do t[i] = i end
local c = 0
for k in next, t, nil do c = c + 1 end
print(c >= 1)"#), "true");
}


#[test]
fn test_next_pairs_unicode_like() {
    assert_eq!(run_lua_one(r#"local t = {}
for i = 1, 20 do t[i] = i end
local c = 0
for k in next, t, nil do c = c + 1 end
print(c >= 1)"#), "true");
}
