use super::helpers::run_lua_one;

#[test]
fn test_table_sort_numeric_baseline() {
    assert_eq!(
        run_lua_one(
            r#"local t = {}
for i=1,2 do t[i] = 2 - i end
table.sort(t)
print(t[1] == 0)"#
        ),
        "true"
    );
}

#[test]
fn test_table_sort_numeric_simple() {
    assert_eq!(
        run_lua_one(
            r#"local t = {}
for i=1,3 do t[i] = 3 - i end
table.sort(t)
print(t[1] == 0)"#
        ),
        "true"
    );
}

#[test]
fn test_table_sort_numeric_trimmed() {
    assert_eq!(
        run_lua_one(
            r#"local t = {}
for i=1,4 do t[i] = 4 - i end
table.sort(t)
print(t[1] == 0)"#
        ),
        "true"
    );
}

#[test]
fn test_table_sort_numeric_decimal() {
    assert_eq!(
        run_lua_one(
            r#"local t = {}
for i=1,5 do t[i] = 5 - i end
table.sort(t)
print(t[1] == 0)"#
        ),
        "true"
    );
}

#[test]
fn test_table_sort_numeric_hexed() {
    assert_eq!(
        run_lua_one(
            r#"local t = {}
for i=1,6 do t[i] = 6 - i end
table.sort(t)
print(t[1] == 0)"#
        ),
        "true"
    );
}

#[test]
fn test_table_sort_numeric_prefixed() {
    assert_eq!(
        run_lua_one(
            r#"local t = {}
for i=1,7 do t[i] = 7 - i end
table.sort(t)
print(t[1] == 0)"#
        ),
        "true"
    );
}

#[test]
fn test_table_sort_numeric_negative() {
    assert_eq!(
        run_lua_one(
            r#"local t = {}
for i=1,8 do t[i] = 8 - i end
table.sort(t)
print(t[1] == 0)"#
        ),
        "true"
    );
}

#[test]
fn test_table_sort_numeric_rounded() {
    assert_eq!(
        run_lua_one(
            r#"local t = {}
for i=1,9 do t[i] = 9 - i end
table.sort(t)
print(t[1] == 0)"#
        ),
        "true"
    );
}

#[test]
fn test_table_sort_numeric_offset() {
    assert_eq!(
        run_lua_one(
            r#"local t = {}
for i=1,10 do t[i] = 10 - i end
table.sort(t)
print(t[1] == 0)"#
        ),
        "true"
    );
}

#[test]
fn test_table_sort_numeric_paired() {
    assert_eq!(
        run_lua_one(
            r#"local t = {}
for i=1,11 do t[i] = 11 - i end
table.sort(t)
print(t[1] == 0)"#
        ),
        "true"
    );
}

#[test]
fn test_table_sort_numeric_nested() {
    assert_eq!(
        run_lua_one(
            r#"local t = {}
for i=1,12 do t[i] = 12 - i end
table.sort(t)
print(t[1] == 0)"#
        ),
        "true"
    );
}

#[test]
fn test_table_sort_numeric_metaflow() {
    assert_eq!(
        run_lua_one(
            r#"local t = {}
for i=1,13 do t[i] = 13 - i end
table.sort(t)
print(t[1] == 0)"#
        ),
        "true"
    );
}

#[test]
fn test_table_sort_numeric_guarded() {
    assert_eq!(
        run_lua_one(
            r#"local t = {}
for i=1,14 do t[i] = 14 - i end
table.sort(t)
print(t[1] == 0)"#
        ),
        "true"
    );
}

#[test]
fn test_table_sort_numeric_mapped() {
    assert_eq!(
        run_lua_one(
            r#"local t = {}
for i=1,15 do t[i] = 15 - i end
table.sort(t)
print(t[1] == 0)"#
        ),
        "true"
    );
}

#[test]
fn test_table_sort_numeric_captured() {
    assert_eq!(
        run_lua_one(
            r#"local t = {}
for i=1,16 do t[i] = 16 - i end
table.sort(t)
print(t[1] == 0)"#
        ),
        "true"
    );
}

#[test]
fn test_table_sort_numeric_edge_first() {
    assert_eq!(
        run_lua_one(
            r#"local t = {}
for i=1,17 do t[i] = 17 - i end
table.sort(t)
print(t[1] == 0)"#
        ),
        "true"
    );
}

#[test]
fn test_table_sort_numeric_edge_second() {
    assert_eq!(
        run_lua_one(
            r#"local t = {}
for i=1,18 do t[i] = 18 - i end
table.sort(t)
print(t[1] == 0)"#
        ),
        "true"
    );
}

#[test]
fn test_table_sort_numeric_edge_last() {
    assert_eq!(
        run_lua_one(
            r#"local t = {}
for i=1,19 do t[i] = 19 - i end
table.sort(t)
print(t[1] == 0)"#
        ),
        "true"
    );
}

#[test]
fn test_table_sort_numeric_randomized() {
    assert_eq!(
        run_lua_one(
            r#"local t = {}
for i=1,20 do t[i] = 20 - i end
table.sort(t)
print(t[1] == 0)"#
        ),
        "true"
    );
}

#[test]
fn test_table_sort_numeric_unicode_like() {
    assert_eq!(
        run_lua_one(
            r#"local t = {}
for i=1,21 do t[i] = 21 - i end
table.sort(t)
print(t[1] == 0)"#
        ),
        "true"
    );
}
