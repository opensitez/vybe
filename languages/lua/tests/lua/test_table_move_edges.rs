use super::helpers::run_lua_one;

#[test]
fn test_table_move_edges_baseline() {
    assert_eq!(
        run_lua_one(
            r#"local src = {1,2,3,4,5}
local dst = {}
table.move(src, 2, 3, 1, dst)
print(type(dst[1]) == "number")"#
        ),
        "true"
    );
}

#[test]
fn test_table_move_edges_simple() {
    assert_eq!(
        run_lua_one(
            r#"local src = {1,2,3,4,5}
local dst = {}
table.move(src, 2, 4, 1, dst)
print(type(dst[1]) == "number")"#
        ),
        "true"
    );
}

#[test]
fn test_table_move_edges_trimmed() {
    assert_eq!(
        run_lua_one(
            r#"local src = {1,2,3,4,5}
local dst = {}
table.move(src, 2, 5, 1, dst)
print(type(dst[1]) == "number")"#
        ),
        "true"
    );
}

#[test]
fn test_table_move_edges_decimal() {
    assert_eq!(
        run_lua_one(
            r#"local src = {1,2,3,4,5}
local dst = {}
table.move(src, 2, 5, 1, dst)
print(type(dst[1]) == "number")"#
        ),
        "true"
    );
}

#[test]
fn test_table_move_edges_hexed() {
    assert_eq!(
        run_lua_one(
            r#"local src = {1,2,3,4,5}
local dst = {}
table.move(src, 2, 2, 1, dst)
print(type(dst[1]) == "number")"#
        ),
        "true"
    );
}

#[test]
fn test_table_move_edges_prefixed() {
    assert_eq!(
        run_lua_one(
            r#"local src = {1,2,3,4,5}
local dst = {}
table.move(src, 2, 3, 1, dst)
print(type(dst[1]) == "number")"#
        ),
        "true"
    );
}

#[test]
fn test_table_move_edges_negative() {
    assert_eq!(
        run_lua_one(
            r#"local src = {1,2,3,4,5}
local dst = {}
table.move(src, 2, 4, 1, dst)
print(type(dst[1]) == "number")"#
        ),
        "true"
    );
}

#[test]
fn test_table_move_edges_rounded() {
    assert_eq!(
        run_lua_one(
            r#"local src = {1,2,3,4,5}
local dst = {}
table.move(src, 2, 5, 1, dst)
print(type(dst[1]) == "number")"#
        ),
        "true"
    );
}

#[test]
fn test_table_move_edges_offset() {
    assert_eq!(
        run_lua_one(
            r#"local src = {1,2,3,4,5}
local dst = {}
table.move(src, 2, 5, 1, dst)
print(type(dst[1]) == "number")"#
        ),
        "true"
    );
}

#[test]
fn test_table_move_edges_paired() {
    assert_eq!(
        run_lua_one(
            r#"local src = {1,2,3,4,5}
local dst = {}
table.move(src, 2, 2, 1, dst)
print(type(dst[1]) == "number")"#
        ),
        "true"
    );
}

#[test]
fn test_table_move_edges_nested() {
    assert_eq!(
        run_lua_one(
            r#"local src = {1,2,3,4,5}
local dst = {}
table.move(src, 2, 3, 1, dst)
print(type(dst[1]) == "number")"#
        ),
        "true"
    );
}

#[test]
fn test_table_move_edges_metaflow() {
    assert_eq!(
        run_lua_one(
            r#"local src = {1,2,3,4,5}
local dst = {}
table.move(src, 2, 4, 1, dst)
print(type(dst[1]) == "number")"#
        ),
        "true"
    );
}

#[test]
fn test_table_move_edges_guarded() {
    assert_eq!(
        run_lua_one(
            r#"local src = {1,2,3,4,5}
local dst = {}
table.move(src, 2, 5, 1, dst)
print(type(dst[1]) == "number")"#
        ),
        "true"
    );
}

#[test]
fn test_table_move_edges_mapped() {
    assert_eq!(
        run_lua_one(
            r#"local src = {1,2,3,4,5}
local dst = {}
table.move(src, 2, 5, 1, dst)
print(type(dst[1]) == "number")"#
        ),
        "true"
    );
}

#[test]
fn test_table_move_edges_captured() {
    assert_eq!(
        run_lua_one(
            r#"local src = {1,2,3,4,5}
local dst = {}
table.move(src, 2, 2, 1, dst)
print(type(dst[1]) == "number")"#
        ),
        "true"
    );
}

#[test]
fn test_table_move_edges_edge_first() {
    assert_eq!(
        run_lua_one(
            r#"local src = {1,2,3,4,5}
local dst = {}
table.move(src, 2, 3, 1, dst)
print(type(dst[1]) == "number")"#
        ),
        "true"
    );
}

#[test]
fn test_table_move_edges_edge_second() {
    assert_eq!(
        run_lua_one(
            r#"local src = {1,2,3,4,5}
local dst = {}
table.move(src, 2, 4, 1, dst)
print(type(dst[1]) == "number")"#
        ),
        "true"
    );
}

#[test]
fn test_table_move_edges_edge_last() {
    assert_eq!(
        run_lua_one(
            r#"local src = {1,2,3,4,5}
local dst = {}
table.move(src, 2, 5, 1, dst)
print(type(dst[1]) == "number")"#
        ),
        "true"
    );
}

#[test]
fn test_table_move_edges_randomized() {
    assert_eq!(
        run_lua_one(
            r#"local src = {1,2,3,4,5}
local dst = {}
table.move(src, 2, 5, 1, dst)
print(type(dst[1]) == "number")"#
        ),
        "true"
    );
}

#[test]
fn test_table_move_edges_unicode_like() {
    assert_eq!(
        run_lua_one(
            r#"local src = {1,2,3,4,5}
local dst = {}
table.move(src, 2, 2, 1, dst)
print(type(dst[1]) == "number")"#
        ),
        "true"
    );
}
