use super::helpers::run_lua_one;

#[test]
fn test_table_insert_edges_baseline() {
    assert_eq!(run_lua_one(r#"local t = {1,2,3}; table.insert(t, 2, 1); print(t[2] == 1)"#), "true");
}


#[test]
fn test_table_insert_edges_simple() {
    assert_eq!(run_lua_one(r#"local t = {1,2,3}; table.insert(t, 3, 2); print(t[3] == 2)"#), "true");
}


#[test]
fn test_table_insert_edges_trimmed() {
    assert_eq!(run_lua_one(r#"local t = {1,2,3}; table.insert(t, 4, 3); print(t[4] == 3)"#), "true");
}


#[test]
fn test_table_insert_edges_decimal() {
    assert_eq!(run_lua_one(r#"local t = {1,2,3}; local ok = pcall(function() table.insert(t, 5, 4) end); print(ok)"#), "false");
}


#[test]
fn test_table_insert_edges_hexed() {
    assert_eq!(run_lua_one(r#"local t = {1,2,3}; local ok = pcall(function() table.insert(t, 5, 4) end); print(ok)"#), "false");
}


#[test]
fn test_table_insert_edges_prefixed() {
    assert_eq!(run_lua_one(r#"local t = {1,2,3}; local ok = pcall(function() table.insert(t, 5, 4) end); print(ok)"#), "false");
}


#[test]
fn test_table_insert_edges_negative() {
    assert_eq!(run_lua_one(r#"local t = {1,2,3}; local ok = pcall(function() table.insert(t, 5, 4) end); print(ok)"#), "false");
}


#[test]
fn test_table_insert_edges_rounded() {
    assert_eq!(run_lua_one(r#"local t = {1,2,3}; local ok = pcall(function() table.insert(t, 5, 4) end); print(ok)"#), "false");
}


#[test]
fn test_table_insert_edges_offset() {
    assert_eq!(run_lua_one(r#"local t = {1,2,3}; local ok = pcall(function() table.insert(t, 5, 4) end); print(ok)"#), "false");
}


#[test]
fn test_table_insert_edges_paired() {
    assert_eq!(run_lua_one(r#"local t = {1,2,3}; local ok = pcall(function() table.insert(t, 5, 4) end); print(ok)"#), "false");
}


#[test]
fn test_table_insert_edges_nested() {
    assert_eq!(run_lua_one(r#"local t = {1,2,3}; local ok = pcall(function() table.insert(t, 5, 4) end); print(ok)"#), "false");
}


#[test]
fn test_table_insert_edges_metaflow() {
    assert_eq!(run_lua_one(r#"local t = {1,2,3}; local ok = pcall(function() table.insert(t, 5, 4) end); print(ok)"#), "false");
}


#[test]
fn test_table_insert_edges_guarded() {
    assert_eq!(run_lua_one(r#"local t = {1,2,3}; local ok = pcall(function() table.insert(t, 5, 4) end); print(ok)"#), "false");
}


#[test]
fn test_table_insert_edges_mapped() {
    assert_eq!(run_lua_one(r#"local t = {1,2,3}; local ok = pcall(function() table.insert(t, 5, 4) end); print(ok)"#), "false");
}


#[test]
fn test_table_insert_edges_captured() {
    assert_eq!(run_lua_one(r#"local t = {1,2,3}; local ok = pcall(function() table.insert(t, 5, 4) end); print(ok)"#), "false");
}


#[test]
fn test_table_insert_edges_edge_first() {
    assert_eq!(run_lua_one(r#"local t = {1,2,3}; local ok = pcall(function() table.insert(t, 5, 4) end); print(ok)"#), "false");
}


#[test]
fn test_table_insert_edges_edge_second() {
    assert_eq!(run_lua_one(r#"local t = {1,2,3}; local ok = pcall(function() table.insert(t, 5, 4) end); print(ok)"#), "false");
}


#[test]
fn test_table_insert_edges_edge_last() {
    assert_eq!(run_lua_one(r#"local t = {1,2,3}; local ok = pcall(function() table.insert(t, 5, 4) end); print(ok)"#), "false");
}


#[test]
fn test_table_insert_edges_randomized() {
    assert_eq!(run_lua_one(r#"local t = {1,2,3}; local ok = pcall(function() table.insert(t, 5, 4) end); print(ok)"#), "false");
}


#[test]
fn test_table_insert_edges_unicode_like() {
    assert_eq!(run_lua_one(r#"local t = {1,2,3}; local ok = pcall(function() table.insert(t, 21, 20) end); print(ok)"#), "false");
}
