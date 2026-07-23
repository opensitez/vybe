use super::helpers::run_lua_one;

#[test]
fn test_pairs_order_baseline() {
    assert_eq!(
        run_lua_one(
            r#"local t = {}
for i = 1, 2 do table.insert(t, i) end
local total = 0
for _, v in ipairs(t) do total = total + v end
print(total == (2 * (2 + 1)) / 2)"#
        ),
        "true"
    );
}

#[test]
fn test_pairs_order_simple() {
    assert_eq!(
        run_lua_one(
            r#"local t = {}
for i = 1, 3 do table.insert(t, i) end
local total = 0
for _, v in ipairs(t) do total = total + v end
print(total == (3 * (3 + 1)) / 2)"#
        ),
        "true"
    );
}

#[test]
fn test_pairs_order_trimmed() {
    assert_eq!(
        run_lua_one(
            r#"local t = {}
for i = 1, 4 do table.insert(t, i) end
local total = 0
for _, v in ipairs(t) do total = total + v end
print(total == (4 * (4 + 1)) / 2)"#
        ),
        "true"
    );
}

#[test]
fn test_pairs_order_decimal() {
    assert_eq!(
        run_lua_one(
            r#"local t = {}
for i = 1, 5 do table.insert(t, i) end
local total = 0
for _, v in ipairs(t) do total = total + v end
print(total == (5 * (5 + 1)) / 2)"#
        ),
        "true"
    );
}

#[test]
fn test_pairs_order_hexed() {
    assert_eq!(
        run_lua_one(
            r#"local t = {}
for i = 1, 6 do table.insert(t, i) end
local total = 0
for _, v in ipairs(t) do total = total + v end
print(total == (6 * (6 + 1)) / 2)"#
        ),
        "true"
    );
}

#[test]
fn test_pairs_order_prefixed() {
    assert_eq!(
        run_lua_one(
            r#"local t = {}
for i = 1, 7 do table.insert(t, i) end
local total = 0
for _, v in ipairs(t) do total = total + v end
print(total == (7 * (7 + 1)) / 2)"#
        ),
        "true"
    );
}

#[test]
fn test_pairs_order_negative() {
    assert_eq!(
        run_lua_one(
            r#"local t = {}
for i = 1, 8 do table.insert(t, i) end
local total = 0
for _, v in ipairs(t) do total = total + v end
print(total == (8 * (8 + 1)) / 2)"#
        ),
        "true"
    );
}

#[test]
fn test_pairs_order_rounded() {
    assert_eq!(
        run_lua_one(
            r#"local t = {}
for i = 1, 9 do table.insert(t, i) end
local total = 0
for _, v in ipairs(t) do total = total + v end
print(total == (9 * (9 + 1)) / 2)"#
        ),
        "true"
    );
}

#[test]
fn test_pairs_order_offset() {
    assert_eq!(
        run_lua_one(
            r#"local t = {}
for i = 1, 10 do table.insert(t, i) end
local total = 0
for _, v in ipairs(t) do total = total + v end
print(total == (10 * (10 + 1)) / 2)"#
        ),
        "true"
    );
}

#[test]
fn test_pairs_order_paired() {
    assert_eq!(
        run_lua_one(
            r#"local t = {}
for i = 1, 11 do table.insert(t, i) end
local total = 0
for _, v in ipairs(t) do total = total + v end
print(total == (11 * (11 + 1)) / 2)"#
        ),
        "true"
    );
}

#[test]
fn test_pairs_order_nested() {
    assert_eq!(
        run_lua_one(
            r#"local t = {}
for i = 1, 12 do table.insert(t, i) end
local total = 0
for _, v in ipairs(t) do total = total + v end
print(total == (12 * (12 + 1)) / 2)"#
        ),
        "true"
    );
}

#[test]
fn test_pairs_order_metaflow() {
    assert_eq!(
        run_lua_one(
            r#"local t = {}
for i = 1, 13 do table.insert(t, i) end
local total = 0
for _, v in ipairs(t) do total = total + v end
print(total == (13 * (13 + 1)) / 2)"#
        ),
        "true"
    );
}

#[test]
fn test_pairs_order_guarded() {
    assert_eq!(
        run_lua_one(
            r#"local t = {}
for i = 1, 14 do table.insert(t, i) end
local total = 0
for _, v in ipairs(t) do total = total + v end
print(total == (14 * (14 + 1)) / 2)"#
        ),
        "true"
    );
}

#[test]
fn test_pairs_order_mapped() {
    assert_eq!(
        run_lua_one(
            r#"local t = {}
for i = 1, 15 do table.insert(t, i) end
local total = 0
for _, v in ipairs(t) do total = total + v end
print(total == (15 * (15 + 1)) / 2)"#
        ),
        "true"
    );
}

#[test]
fn test_pairs_order_captured() {
    assert_eq!(
        run_lua_one(
            r#"local t = {}
for i = 1, 16 do table.insert(t, i) end
local total = 0
for _, v in ipairs(t) do total = total + v end
print(total == (16 * (16 + 1)) / 2)"#
        ),
        "true"
    );
}

#[test]
fn test_pairs_order_edge_first() {
    assert_eq!(
        run_lua_one(
            r#"local t = {}
for i = 1, 17 do table.insert(t, i) end
local total = 0
for _, v in ipairs(t) do total = total + v end
print(total == (17 * (17 + 1)) / 2)"#
        ),
        "true"
    );
}

#[test]
fn test_pairs_order_edge_second() {
    assert_eq!(
        run_lua_one(
            r#"local t = {}
for i = 1, 18 do table.insert(t, i) end
local total = 0
for _, v in ipairs(t) do total = total + v end
print(total == (18 * (18 + 1)) / 2)"#
        ),
        "true"
    );
}

#[test]
fn test_pairs_order_edge_last() {
    assert_eq!(
        run_lua_one(
            r#"local t = {}
for i = 1, 19 do table.insert(t, i) end
local total = 0
for _, v in ipairs(t) do total = total + v end
print(total == (19 * (19 + 1)) / 2)"#
        ),
        "true"
    );
}

#[test]
fn test_pairs_order_randomized() {
    assert_eq!(
        run_lua_one(
            r#"local t = {}
for i = 1, 20 do table.insert(t, i) end
local total = 0
for _, v in ipairs(t) do total = total + v end
print(total == (20 * (20 + 1)) / 2)"#
        ),
        "true"
    );
}

#[test]
fn test_pairs_order_unicode_like() {
    assert_eq!(
        run_lua_one(
            r#"local t = {}
for i = 1, 21 do table.insert(t, i) end
local total = 0
for _, v in ipairs(t) do total = total + v end
print(total == (21 * (21 + 1)) / 2)"#
        ),
        "true"
    );
}
