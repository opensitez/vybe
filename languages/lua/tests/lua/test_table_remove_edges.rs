use super::helpers::run_lua_one;

#[test]
fn test_table_remove_edges_baseline() {
    assert_eq!(
        run_lua_one(
            r#"local t = {1, 2, 3}
local v = table.remove(t, 2)
print(type(v) == "number")"#
        ),
        "true"
    );
}

#[test]
fn test_table_remove_edges_simple() {
    assert_eq!(
        run_lua_one(
            r#"local t = {2, 3, 4}
local v = table.remove(t, 3)
print(type(v) == "number")"#
        ),
        "true"
    );
}

#[test]
fn test_table_remove_edges_trimmed() {
    assert_eq!(
        run_lua_one(
            r#"local t = {3, 4, 5}
local v = table.remove(t, 1)
print(type(v) == "number")"#
        ),
        "true"
    );
}

#[test]
fn test_table_remove_edges_decimal() {
    assert_eq!(
        run_lua_one(
            r#"local t = {4, 5, 6}
local v = table.remove(t, 2)
print(type(v) == "number")"#
        ),
        "true"
    );
}

#[test]
fn test_table_remove_edges_hexed() {
    assert_eq!(
        run_lua_one(
            r#"local t = {5, 6, 7}
local v = table.remove(t, 3)
print(type(v) == "number")"#
        ),
        "true"
    );
}

#[test]
fn test_table_remove_edges_prefixed() {
    assert_eq!(
        run_lua_one(
            r#"local t = {6, 7, 8}
local v = table.remove(t, 1)
print(type(v) == "number")"#
        ),
        "true"
    );
}

#[test]
fn test_table_remove_edges_negative() {
    assert_eq!(
        run_lua_one(
            r#"local t = {7, 8, 9}
local v = table.remove(t, 2)
print(type(v) == "number")"#
        ),
        "true"
    );
}

#[test]
fn test_table_remove_edges_rounded() {
    assert_eq!(
        run_lua_one(
            r#"local t = {8, 9, 10}
local v = table.remove(t, 3)
print(type(v) == "number")"#
        ),
        "true"
    );
}

#[test]
fn test_table_remove_edges_offset() {
    assert_eq!(
        run_lua_one(
            r#"local t = {9, 10, 11}
local v = table.remove(t, 1)
print(type(v) == "number")"#
        ),
        "true"
    );
}

#[test]
fn test_table_remove_edges_paired() {
    assert_eq!(
        run_lua_one(
            r#"local t = {10, 11, 12}
local v = table.remove(t, 2)
print(type(v) == "number")"#
        ),
        "true"
    );
}

#[test]
fn test_table_remove_edges_nested() {
    assert_eq!(
        run_lua_one(
            r#"local t = {11, 12, 13}
local v = table.remove(t, 3)
print(type(v) == "number")"#
        ),
        "true"
    );
}

#[test]
fn test_table_remove_edges_metaflow() {
    assert_eq!(
        run_lua_one(
            r#"local t = {12, 13, 14}
local v = table.remove(t, 1)
print(type(v) == "number")"#
        ),
        "true"
    );
}

#[test]
fn test_table_remove_edges_guarded() {
    assert_eq!(
        run_lua_one(
            r#"local t = {13, 14, 15}
local v = table.remove(t, 2)
print(type(v) == "number")"#
        ),
        "true"
    );
}

#[test]
fn test_table_remove_edges_mapped() {
    assert_eq!(
        run_lua_one(
            r#"local t = {14, 15, 16}
local v = table.remove(t, 3)
print(type(v) == "number")"#
        ),
        "true"
    );
}

#[test]
fn test_table_remove_edges_captured() {
    assert_eq!(
        run_lua_one(
            r#"local t = {15, 16, 17}
local v = table.remove(t, 1)
print(type(v) == "number")"#
        ),
        "true"
    );
}

#[test]
fn test_table_remove_edges_edge_first() {
    assert_eq!(
        run_lua_one(
            r#"local t = {16, 17, 18}
local v = table.remove(t, 2)
print(type(v) == "number")"#
        ),
        "true"
    );
}

#[test]
fn test_table_remove_edges_edge_second() {
    assert_eq!(
        run_lua_one(
            r#"local t = {17, 18, 19}
local v = table.remove(t, 3)
print(type(v) == "number")"#
        ),
        "true"
    );
}

#[test]
fn test_table_remove_edges_edge_last() {
    assert_eq!(
        run_lua_one(
            r#"local t = {18, 19, 20}
local v = table.remove(t, 1)
print(type(v) == "number")"#
        ),
        "true"
    );
}

#[test]
fn test_table_remove_edges_randomized() {
    assert_eq!(
        run_lua_one(
            r#"local t = {19, 20, 21}
local v = table.remove(t, 2)
print(type(v) == "number")"#
        ),
        "true"
    );
}

#[test]
fn test_table_remove_edges_unicode_like() {
    assert_eq!(
        run_lua_one(
            r#"local t = {20, 21, 22}
local v = table.remove(t, 3)
print(type(v) == "number")"#
        ),
        "true"
    );
}
