use super::helpers::run_lua_one;

#[test]
fn test_table_sort_nested_baseline() {
    assert_eq!(run_lua_one(r#"local t = {{v=3}, {v=1}, {v=2}}
table.sort(t, function(a,b) return a.v < b.v end)
print(t[1].v == 1)"#), "true");
}


#[test]
fn test_table_sort_nested_simple() {
    assert_eq!(run_lua_one(r#"local t = {{v=4}, {v=2}, {v=3}}
table.sort(t, function(a,b) return a.v < b.v end)
print(t[1].v == 2)"#), "true");
}


#[test]
fn test_table_sort_nested_trimmed() {
    assert_eq!(run_lua_one(r#"local t = {{v=5}, {v=3}, {v=4}}
table.sort(t, function(a,b) return a.v < b.v end)
print(t[1].v == 3)"#), "true");
}


#[test]
fn test_table_sort_nested_decimal() {
    assert_eq!(run_lua_one(r#"local t = {{v=6}, {v=4}, {v=5}}
table.sort(t, function(a,b) return a.v < b.v end)
print(t[1].v == 4)"#), "true");
}


#[test]
fn test_table_sort_nested_hexed() {
    assert_eq!(run_lua_one(r#"local t = {{v=7}, {v=5}, {v=6}}
table.sort(t, function(a,b) return a.v < b.v end)
print(t[1].v == 5)"#), "true");
}


#[test]
fn test_table_sort_nested_prefixed() {
    assert_eq!(run_lua_one(r#"local t = {{v=8}, {v=6}, {v=7}}
table.sort(t, function(a,b) return a.v < b.v end)
print(t[1].v == 6)"#), "true");
}


#[test]
fn test_table_sort_nested_negative() {
    assert_eq!(run_lua_one(r#"local t = {{v=9}, {v=7}, {v=8}}
table.sort(t, function(a,b) return a.v < b.v end)
print(t[1].v == 7)"#), "true");
}


#[test]
fn test_table_sort_nested_rounded() {
    assert_eq!(run_lua_one(r#"local t = {{v=10}, {v=8}, {v=9}}
table.sort(t, function(a,b) return a.v < b.v end)
print(t[1].v == 8)"#), "true");
}


#[test]
fn test_table_sort_nested_offset() {
    assert_eq!(run_lua_one(r#"local t = {{v=11}, {v=9}, {v=10}}
table.sort(t, function(a,b) return a.v < b.v end)
print(t[1].v == 9)"#), "true");
}


#[test]
fn test_table_sort_nested_paired() {
    assert_eq!(run_lua_one(r#"local t = {{v=12}, {v=10}, {v=11}}
table.sort(t, function(a,b) return a.v < b.v end)
print(t[1].v == 10)"#), "true");
}


#[test]
fn test_table_sort_nested_nested() {
    assert_eq!(run_lua_one(r#"local t = {{v=13}, {v=11}, {v=12}}
table.sort(t, function(a,b) return a.v < b.v end)
print(t[1].v == 11)"#), "true");
}


#[test]
fn test_table_sort_nested_metaflow() {
    assert_eq!(run_lua_one(r#"local t = {{v=14}, {v=12}, {v=13}}
table.sort(t, function(a,b) return a.v < b.v end)
print(t[1].v == 12)"#), "true");
}


#[test]
fn test_table_sort_nested_guarded() {
    assert_eq!(run_lua_one(r#"local t = {{v=15}, {v=13}, {v=14}}
table.sort(t, function(a,b) return a.v < b.v end)
print(t[1].v == 13)"#), "true");
}


#[test]
fn test_table_sort_nested_mapped() {
    assert_eq!(run_lua_one(r#"local t = {{v=16}, {v=14}, {v=15}}
table.sort(t, function(a,b) return a.v < b.v end)
print(t[1].v == 14)"#), "true");
}


#[test]
fn test_table_sort_nested_captured() {
    assert_eq!(run_lua_one(r#"local t = {{v=17}, {v=15}, {v=16}}
table.sort(t, function(a,b) return a.v < b.v end)
print(t[1].v == 15)"#), "true");
}


#[test]
fn test_table_sort_nested_edge_first() {
    assert_eq!(run_lua_one(r#"local t = {{v=18}, {v=16}, {v=17}}
table.sort(t, function(a,b) return a.v < b.v end)
print(t[1].v == 16)"#), "true");
}


#[test]
fn test_table_sort_nested_edge_second() {
    assert_eq!(run_lua_one(r#"local t = {{v=19}, {v=17}, {v=18}}
table.sort(t, function(a,b) return a.v < b.v end)
print(t[1].v == 17)"#), "true");
}


#[test]
fn test_table_sort_nested_edge_last() {
    assert_eq!(run_lua_one(r#"local t = {{v=20}, {v=18}, {v=19}}
table.sort(t, function(a,b) return a.v < b.v end)
print(t[1].v == 18)"#), "true");
}


#[test]
fn test_table_sort_nested_randomized() {
    assert_eq!(run_lua_one(r#"local t = {{v=21}, {v=19}, {v=20}}
table.sort(t, function(a,b) return a.v < b.v end)
print(t[1].v == 19)"#), "true");
}


#[test]
fn test_table_sort_nested_unicode_like() {
    assert_eq!(run_lua_one(r#"local t = {{v=22}, {v=20}, {v=21}}
table.sort(t, function(a,b) return a.v < b.v end)
print(t[1].v == 20)"#), "true");
}
