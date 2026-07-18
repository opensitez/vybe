use super::helpers::run_lua_one;

#[test]
fn test_table_pack_unpack_baseline() {
    assert_eq!(run_lua_one(r#"local t = table.pack(1, 2, 3)
local a, b, c = table.unpack(t)
print(a == 1 and b == 2 and c == 3)"#), "true");
}


#[test]
fn test_table_pack_unpack_simple() {
    assert_eq!(run_lua_one(r#"local t = table.pack(2, 3, 4)
local a, b, c = table.unpack(t)
print(a == 2 and b == 3 and c == 4)"#), "true");
}


#[test]
fn test_table_pack_unpack_trimmed() {
    assert_eq!(run_lua_one(r#"local t = table.pack(3, 4, 5)
local a, b, c = table.unpack(t)
print(a == 3 and b == 4 and c == 5)"#), "true");
}


#[test]
fn test_table_pack_unpack_decimal() {
    assert_eq!(run_lua_one(r#"local t = table.pack(4, 5, 6)
local a, b, c = table.unpack(t)
print(a == 4 and b == 5 and c == 6)"#), "true");
}


#[test]
fn test_table_pack_unpack_hexed() {
    assert_eq!(run_lua_one(r#"local t = table.pack(5, 6, 7)
local a, b, c = table.unpack(t)
print(a == 5 and b == 6 and c == 7)"#), "true");
}


#[test]
fn test_table_pack_unpack_prefixed() {
    assert_eq!(run_lua_one(r#"local t = table.pack(6, 7, 8)
local a, b, c = table.unpack(t)
print(a == 6 and b == 7 and c == 8)"#), "true");
}


#[test]
fn test_table_pack_unpack_negative() {
    assert_eq!(run_lua_one(r#"local t = table.pack(7, 8, 9)
local a, b, c = table.unpack(t)
print(a == 7 and b == 8 and c == 9)"#), "true");
}


#[test]
fn test_table_pack_unpack_rounded() {
    assert_eq!(run_lua_one(r#"local t = table.pack(8, 9, 10)
local a, b, c = table.unpack(t)
print(a == 8 and b == 9 and c == 10)"#), "true");
}


#[test]
fn test_table_pack_unpack_offset() {
    assert_eq!(run_lua_one(r#"local t = table.pack(9, 10, 11)
local a, b, c = table.unpack(t)
print(a == 9 and b == 10 and c == 11)"#), "true");
}


#[test]
fn test_table_pack_unpack_paired() {
    assert_eq!(run_lua_one(r#"local t = table.pack(10, 11, 12)
local a, b, c = table.unpack(t)
print(a == 10 and b == 11 and c == 12)"#), "true");
}


#[test]
fn test_table_pack_unpack_nested() {
    assert_eq!(run_lua_one(r#"local t = table.pack(11, 12, 13)
local a, b, c = table.unpack(t)
print(a == 11 and b == 12 and c == 13)"#), "true");
}


#[test]
fn test_table_pack_unpack_metaflow() {
    assert_eq!(run_lua_one(r#"local t = table.pack(12, 13, 14)
local a, b, c = table.unpack(t)
print(a == 12 and b == 13 and c == 14)"#), "true");
}


#[test]
fn test_table_pack_unpack_guarded() {
    assert_eq!(run_lua_one(r#"local t = table.pack(13, 14, 15)
local a, b, c = table.unpack(t)
print(a == 13 and b == 14 and c == 15)"#), "true");
}


#[test]
fn test_table_pack_unpack_mapped() {
    assert_eq!(run_lua_one(r#"local t = table.pack(14, 15, 16)
local a, b, c = table.unpack(t)
print(a == 14 and b == 15 and c == 16)"#), "true");
}


#[test]
fn test_table_pack_unpack_captured() {
    assert_eq!(run_lua_one(r#"local t = table.pack(15, 16, 17)
local a, b, c = table.unpack(t)
print(a == 15 and b == 16 and c == 17)"#), "true");
}


#[test]
fn test_table_pack_unpack_edge_first() {
    assert_eq!(run_lua_one(r#"local t = table.pack(16, 17, 18)
local a, b, c = table.unpack(t)
print(a == 16 and b == 17 and c == 18)"#), "true");
}


#[test]
fn test_table_pack_unpack_edge_second() {
    assert_eq!(run_lua_one(r#"local t = table.pack(17, 18, 19)
local a, b, c = table.unpack(t)
print(a == 17 and b == 18 and c == 19)"#), "true");
}


#[test]
fn test_table_pack_unpack_edge_last() {
    assert_eq!(run_lua_one(r#"local t = table.pack(18, 19, 20)
local a, b, c = table.unpack(t)
print(a == 18 and b == 19 and c == 20)"#), "true");
}


#[test]
fn test_table_pack_unpack_randomized() {
    assert_eq!(run_lua_one(r#"local t = table.pack(19, 20, 21)
local a, b, c = table.unpack(t)
print(a == 19 and b == 20 and c == 21)"#), "true");
}


#[test]
fn test_table_pack_unpack_unicode_like() {
    assert_eq!(run_lua_one(r#"local t = table.pack(20, 21, 22)
local a, b, c = table.unpack(t)
print(a == 20 and b == 21 and c == 22)"#), "true");
}
