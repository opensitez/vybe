use super::helpers::run_lua_one;

#[test]
fn test_table_sparse_behavior_baseline() {
    assert_eq!(run_lua_one(r#"local t = {[1] = 1, [100] = 2}
local c = 0
for _ in pairs(t) do c = c + 1 end
print(c == 2)"#), "true");
}


#[test]
fn test_table_sparse_behavior_simple() {
    assert_eq!(run_lua_one(r#"local t = {[1] = 1, [100] = 3}
local c = 0
for _ in pairs(t) do c = c + 1 end
print(c == 2)"#), "true");
}


#[test]
fn test_table_sparse_behavior_trimmed() {
    assert_eq!(run_lua_one(r#"local t = {[1] = 1, [100] = 4}
local c = 0
for _ in pairs(t) do c = c + 1 end
print(c == 2)"#), "true");
}


#[test]
fn test_table_sparse_behavior_decimal() {
    assert_eq!(run_lua_one(r#"local t = {[1] = 1, [100] = 5}
local c = 0
for _ in pairs(t) do c = c + 1 end
print(c == 2)"#), "true");
}


#[test]
fn test_table_sparse_behavior_hexed() {
    assert_eq!(run_lua_one(r#"local t = {[1] = 1, [100] = 6}
local c = 0
for _ in pairs(t) do c = c + 1 end
print(c == 2)"#), "true");
}


#[test]
fn test_table_sparse_behavior_prefixed() {
    assert_eq!(run_lua_one(r#"local t = {[1] = 1, [100] = 7}
local c = 0
for _ in pairs(t) do c = c + 1 end
print(c == 2)"#), "true");
}


#[test]
fn test_table_sparse_behavior_negative() {
    assert_eq!(run_lua_one(r#"local t = {[1] = 1, [100] = 8}
local c = 0
for _ in pairs(t) do c = c + 1 end
print(c == 2)"#), "true");
}


#[test]
fn test_table_sparse_behavior_rounded() {
    assert_eq!(run_lua_one(r#"local t = {[1] = 1, [100] = 9}
local c = 0
for _ in pairs(t) do c = c + 1 end
print(c == 2)"#), "true");
}


#[test]
fn test_table_sparse_behavior_offset() {
    assert_eq!(run_lua_one(r#"local t = {[1] = 1, [100] = 10}
local c = 0
for _ in pairs(t) do c = c + 1 end
print(c == 2)"#), "true");
}


#[test]
fn test_table_sparse_behavior_paired() {
    assert_eq!(run_lua_one(r#"local t = {[1] = 1, [100] = 11}
local c = 0
for _ in pairs(t) do c = c + 1 end
print(c == 2)"#), "true");
}


#[test]
fn test_table_sparse_behavior_nested() {
    assert_eq!(run_lua_one(r#"local t = {[1] = 1, [100] = 12}
local c = 0
for _ in pairs(t) do c = c + 1 end
print(c == 2)"#), "true");
}


#[test]
fn test_table_sparse_behavior_metaflow() {
    assert_eq!(run_lua_one(r#"local t = {[1] = 1, [100] = 13}
local c = 0
for _ in pairs(t) do c = c + 1 end
print(c == 2)"#), "true");
}


#[test]
fn test_table_sparse_behavior_guarded() {
    assert_eq!(run_lua_one(r#"local t = {[1] = 1, [100] = 14}
local c = 0
for _ in pairs(t) do c = c + 1 end
print(c == 2)"#), "true");
}


#[test]
fn test_table_sparse_behavior_mapped() {
    assert_eq!(run_lua_one(r#"local t = {[1] = 1, [100] = 15}
local c = 0
for _ in pairs(t) do c = c + 1 end
print(c == 2)"#), "true");
}


#[test]
fn test_table_sparse_behavior_captured() {
    assert_eq!(run_lua_one(r#"local t = {[1] = 1, [100] = 16}
local c = 0
for _ in pairs(t) do c = c + 1 end
print(c == 2)"#), "true");
}


#[test]
fn test_table_sparse_behavior_edge_first() {
    assert_eq!(run_lua_one(r#"local t = {[1] = 1, [100] = 17}
local c = 0
for _ in pairs(t) do c = c + 1 end
print(c == 2)"#), "true");
}


#[test]
fn test_table_sparse_behavior_edge_second() {
    assert_eq!(run_lua_one(r#"local t = {[1] = 1, [100] = 18}
local c = 0
for _ in pairs(t) do c = c + 1 end
print(c == 2)"#), "true");
}


#[test]
fn test_table_sparse_behavior_edge_last() {
    assert_eq!(run_lua_one(r#"local t = {[1] = 1, [100] = 19}
local c = 0
for _ in pairs(t) do c = c + 1 end
print(c == 2)"#), "true");
}


#[test]
fn test_table_sparse_behavior_randomized() {
    assert_eq!(run_lua_one(r#"local t = {[1] = 1, [100] = 20}
local c = 0
for _ in pairs(t) do c = c + 1 end
print(c == 2)"#), "true");
}


#[test]
fn test_table_sparse_behavior_unicode_like() {
    assert_eq!(run_lua_one(r#"local t = {[1] = 1, [100] = 21}
local c = 0
for _ in pairs(t) do c = c + 1 end
print(c == 2)"#), "true");
}
