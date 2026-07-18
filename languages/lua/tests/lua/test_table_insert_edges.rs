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
    assert_eq!(run_lua_one(r#"local t = {1,2,3}; table.insert(t, 5, 4); print(t[5] == 4)"#), "true");
}


#[test]
fn test_table_insert_edges_hexed() {
    assert_eq!(run_lua_one(r#"local t = {1,2,3}; table.insert(t, 6, 5); print(t[6] == 5)"#), "true");
}


#[test]
fn test_table_insert_edges_prefixed() {
    assert_eq!(run_lua_one(r#"local t = {1,2,3}; table.insert(t, 7, 6); print(t[7] == 6)"#), "true");
}


#[test]
fn test_table_insert_edges_negative() {
    assert_eq!(run_lua_one(r#"local t = {1,2,3}; table.insert(t, 8, 7); print(t[8] == 7)"#), "true");
}


#[test]
fn test_table_insert_edges_rounded() {
    assert_eq!(run_lua_one(r#"local t = {1,2,3}; table.insert(t, 9, 8); print(t[9] == 8)"#), "true");
}


#[test]
fn test_table_insert_edges_offset() {
    assert_eq!(run_lua_one(r#"local t = {1,2,3}; table.insert(t, 10, 9); print(t[10] == 9)"#), "true");
}


#[test]
fn test_table_insert_edges_paired() {
    assert_eq!(run_lua_one(r#"local t = {1,2,3}; table.insert(t, 11, 10); print(t[11] == 10)"#), "true");
}


#[test]
fn test_table_insert_edges_nested() {
    assert_eq!(run_lua_one(r#"local t = {1,2,3}; table.insert(t, 12, 11); print(t[12] == 11)"#), "true");
}


#[test]
fn test_table_insert_edges_metaflow() {
    assert_eq!(run_lua_one(r#"local t = {1,2,3}; table.insert(t, 13, 12); print(t[13] == 12)"#), "true");
}


#[test]
fn test_table_insert_edges_guarded() {
    assert_eq!(run_lua_one(r#"local t = {1,2,3}; table.insert(t, 14, 13); print(t[14] == 13)"#), "true");
}


#[test]
fn test_table_insert_edges_mapped() {
    assert_eq!(run_lua_one(r#"local t = {1,2,3}; table.insert(t, 15, 14); print(t[15] == 14)"#), "true");
}


#[test]
fn test_table_insert_edges_captured() {
    assert_eq!(run_lua_one(r#"local t = {1,2,3}; table.insert(t, 16, 15); print(t[16] == 15)"#), "true");
}


#[test]
fn test_table_insert_edges_edge_first() {
    assert_eq!(run_lua_one(r#"local t = {1,2,3}; table.insert(t, 17, 16); print(t[17] == 16)"#), "true");
}


#[test]
fn test_table_insert_edges_edge_second() {
    assert_eq!(run_lua_one(r#"local t = {1,2,3}; table.insert(t, 18, 17); print(t[18] == 17)"#), "true");
}


#[test]
fn test_table_insert_edges_edge_last() {
    assert_eq!(run_lua_one(r#"local t = {1,2,3}; table.insert(t, 19, 18); print(t[19] == 18)"#), "true");
}


#[test]
fn test_table_insert_edges_randomized() {
    assert_eq!(run_lua_one(r#"local t = {1,2,3}; table.insert(t, 20, 19); print(t[20] == 19)"#), "true");
}


#[test]
fn test_table_insert_edges_unicode_like() {
    assert_eq!(run_lua_one(r#"local t = {1,2,3}; table.insert(t, 21, 20); print(t[21] == 20)"#), "true");
}
