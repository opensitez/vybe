use super::helpers::run_lua_one;

#[test]
fn test_table_concat_edge_baseline() {
    assert_eq!(run_lua_one(r#"local t = {1, 2, 3}
print(table.concat(t, ",") == ("1,2,3"))"#), "true");
}


#[test]
fn test_table_concat_edge_simple() {
    assert_eq!(run_lua_one(r#"local t = {2, 3, 5}
print(table.concat(t, ",") == ("2,3,5"))"#), "true");
}


#[test]
fn test_table_concat_edge_trimmed() {
    assert_eq!(run_lua_one(r#"local t = {3, 4, 7}
print(table.concat(t, ",") == ("3,4,7"))"#), "true");
}


#[test]
fn test_table_concat_edge_decimal() {
    assert_eq!(run_lua_one(r#"local t = {4, 5, 9}
print(table.concat(t, ",") == ("4,5,9"))"#), "true");
}


#[test]
fn test_table_concat_edge_hexed() {
    assert_eq!(run_lua_one(r#"local t = {5, 6, 11}
print(table.concat(t, ",") == ("5,6,11"))"#), "true");
}


#[test]
fn test_table_concat_edge_prefixed() {
    assert_eq!(run_lua_one(r#"local t = {6, 7, 13}
print(table.concat(t, ",") == ("6,7,13"))"#), "true");
}


#[test]
fn test_table_concat_edge_negative() {
    assert_eq!(run_lua_one(r#"local t = {7, 8, 15}
print(table.concat(t, ",") == ("7,8,15"))"#), "true");
}


#[test]
fn test_table_concat_edge_rounded() {
    assert_eq!(run_lua_one(r#"local t = {8, 9, 17}
print(table.concat(t, ",") == ("8,9,17"))"#), "true");
}


#[test]
fn test_table_concat_edge_offset() {
    assert_eq!(run_lua_one(r#"local t = {9, 10, 19}
print(table.concat(t, ",") == ("9,10,19"))"#), "true");
}


#[test]
fn test_table_concat_edge_paired() {
    assert_eq!(run_lua_one(r#"local t = {10, 11, 21}
print(table.concat(t, ",") == ("10,11,21"))"#), "true");
}


#[test]
fn test_table_concat_edge_nested() {
    assert_eq!(run_lua_one(r#"local t = {11, 12, 23}
print(table.concat(t, ",") == ("11,12,23"))"#), "true");
}


#[test]
fn test_table_concat_edge_metaflow() {
    assert_eq!(run_lua_one(r#"local t = {12, 13, 25}
print(table.concat(t, ",") == ("12,13,25"))"#), "true");
}


#[test]
fn test_table_concat_edge_guarded() {
    assert_eq!(run_lua_one(r#"local t = {13, 14, 27}
print(table.concat(t, ",") == ("13,14,27"))"#), "true");
}


#[test]
fn test_table_concat_edge_mapped() {
    assert_eq!(run_lua_one(r#"local t = {14, 15, 29}
print(table.concat(t, ",") == ("14,15,29"))"#), "true");
}


#[test]
fn test_table_concat_edge_captured() {
    assert_eq!(run_lua_one(r#"local t = {15, 16, 31}
print(table.concat(t, ",") == ("15,16,31"))"#), "true");
}


#[test]
fn test_table_concat_edge_edge_first() {
    assert_eq!(run_lua_one(r#"local t = {16, 17, 33}
print(table.concat(t, ",") == ("16,17,33"))"#), "true");
}


#[test]
fn test_table_concat_edge_edge_second() {
    assert_eq!(run_lua_one(r#"local t = {17, 18, 35}
print(table.concat(t, ",") == ("17,18,35"))"#), "true");
}


#[test]
fn test_table_concat_edge_edge_last() {
    assert_eq!(run_lua_one(r#"local t = {18, 19, 37}
print(table.concat(t, ",") == ("18,19,37"))"#), "true");
}


#[test]
fn test_table_concat_edge_randomized() {
    assert_eq!(run_lua_one(r#"local t = {19, 20, 39}
print(table.concat(t, ",") == ("19,20,39"))"#), "true");
}


#[test]
fn test_table_concat_edge_unicode_like() {
    assert_eq!(run_lua_one(r#"local t = {20, 21, 41}
print(table.concat(t, ",") == ("20,21,41"))"#), "true");
}
