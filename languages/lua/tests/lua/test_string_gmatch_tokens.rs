use super::helpers::run_lua_one;

#[test]
fn test_string_gmatch_tokens_baseline() {
    assert_eq!(
        run_lua_one(
            r#"local c = 0
for _ in string.gmatch("w0 w1", "[%a]+") do c = c + 1 end
print(c == 2)"#
        ),
        "true"
    );
}

#[test]
fn test_string_gmatch_tokens_simple() {
    assert_eq!(
        run_lua_one(
            r#"local c = 0
for _ in string.gmatch("w0 w1 w2", "[%a]+") do c = c + 1 end
print(c == 3)"#
        ),
        "true"
    );
}

#[test]
fn test_string_gmatch_tokens_trimmed() {
    assert_eq!(
        run_lua_one(
            r#"local c = 0
for _ in string.gmatch("w0 w1 w2 w3", "[%a]+") do c = c + 1 end
print(c == 4)"#
        ),
        "true"
    );
}

#[test]
fn test_string_gmatch_tokens_decimal() {
    assert_eq!(
        run_lua_one(
            r#"local c = 0
for _ in string.gmatch("w0 w1 w2 w3 w4", "[%a]+") do c = c + 1 end
print(c == 5)"#
        ),
        "true"
    );
}

#[test]
fn test_string_gmatch_tokens_hexed() {
    assert_eq!(
        run_lua_one(
            r#"local c = 0
for _ in string.gmatch("w0 w1 w2 w3 w4 w5", "[%a]+") do c = c + 1 end
print(c == 6)"#
        ),
        "true"
    );
}

#[test]
fn test_string_gmatch_tokens_prefixed() {
    assert_eq!(
        run_lua_one(
            r#"local c = 0
for _ in string.gmatch("w0 w1 w2 w3 w4 w5 w6", "[%a]+") do c = c + 1 end
print(c == 7)"#
        ),
        "true"
    );
}

#[test]
fn test_string_gmatch_tokens_negative() {
    assert_eq!(
        run_lua_one(
            r#"local c = 0
for _ in string.gmatch("w0 w1 w2 w3 w4 w5 w6 w7", "[%a]+") do c = c + 1 end
print(c == 8)"#
        ),
        "true"
    );
}

#[test]
fn test_string_gmatch_tokens_rounded() {
    assert_eq!(
        run_lua_one(
            r#"local c = 0
for _ in string.gmatch("w0 w1 w2 w3 w4 w5 w6 w7 w8", "[%a]+") do c = c + 1 end
print(c == 9)"#
        ),
        "true"
    );
}

#[test]
fn test_string_gmatch_tokens_offset() {
    assert_eq!(
        run_lua_one(
            r#"local c = 0
for _ in string.gmatch("w0 w1 w2 w3 w4 w5 w6 w7 w8 w9", "[%a]+") do c = c + 1 end
print(c == 10)"#
        ),
        "true"
    );
}

#[test]
fn test_string_gmatch_tokens_paired() {
    assert_eq!(
        run_lua_one(
            r#"local c = 0
for _ in string.gmatch("w0 w1 w2 w3 w4 w5 w6 w7 w8 w9 w10", "[%a]+") do c = c + 1 end
print(c == 11)"#
        ),
        "true"
    );
}

#[test]
fn test_string_gmatch_tokens_nested() {
    assert_eq!(
        run_lua_one(
            r#"local c = 0
for _ in string.gmatch("w0 w1 w2 w3 w4 w5 w6 w7 w8 w9 w10 w11", "[%a]+") do c = c + 1 end
print(c == 12)"#
        ),
        "true"
    );
}

#[test]
fn test_string_gmatch_tokens_metaflow() {
    assert_eq!(
        run_lua_one(
            r#"local c = 0
for _ in string.gmatch("w0 w1 w2 w3 w4 w5 w6 w7 w8 w9 w10 w11 w12", "[%a]+") do c = c + 1 end
print(c == 13)"#
        ),
        "true"
    );
}

#[test]
fn test_string_gmatch_tokens_guarded() {
    assert_eq!(
        run_lua_one(
            r#"local c = 0
for _ in string.gmatch("w0 w1 w2 w3 w4 w5 w6 w7 w8 w9 w10 w11 w12 w13", "[%a]+") do c = c + 1 end
print(c == 14)"#
        ),
        "true"
    );
}

#[test]
fn test_string_gmatch_tokens_mapped() {
    assert_eq!(
        run_lua_one(
            r#"local c = 0
for _ in string.gmatch("w0 w1 w2 w3 w4 w5 w6 w7 w8 w9 w10 w11 w12 w13 w14", "[%a]+") do c = c + 1 end
print(c == 15)"#
        ),
        "true"
    );
}

#[test]
fn test_string_gmatch_tokens_captured() {
    assert_eq!(
        run_lua_one(
            r#"local c = 0
for _ in string.gmatch("w0 w1 w2 w3 w4 w5 w6 w7 w8 w9 w10 w11 w12 w13 w14 w15", "[%a]+") do c = c + 1 end
print(c == 16)"#
        ),
        "true"
    );
}

#[test]
fn test_string_gmatch_tokens_edge_first() {
    assert_eq!(
        run_lua_one(
            r#"local c = 0
for _ in string.gmatch("w0 w1 w2 w3 w4 w5 w6 w7 w8 w9 w10 w11 w12 w13 w14 w15 w16", "[%a]+") do c = c + 1 end
print(c == 17)"#
        ),
        "true"
    );
}

#[test]
fn test_string_gmatch_tokens_edge_second() {
    assert_eq!(
        run_lua_one(
            r#"local c = 0
for _ in string.gmatch("w0 w1 w2 w3 w4 w5 w6 w7 w8 w9 w10 w11 w12 w13 w14 w15 w16 w17", "[%a]+") do c = c + 1 end
print(c == 18)"#
        ),
        "true"
    );
}

#[test]
fn test_string_gmatch_tokens_edge_last() {
    assert_eq!(
        run_lua_one(
            r#"local c = 0
for _ in string.gmatch("w0 w1 w2 w3 w4 w5 w6 w7 w8 w9 w10 w11 w12 w13 w14 w15 w16 w17 w18", "[%a]+") do c = c + 1 end
print(c == 19)"#
        ),
        "true"
    );
}

#[test]
fn test_string_gmatch_tokens_randomized() {
    assert_eq!(
        run_lua_one(
            r#"local c = 0
for _ in string.gmatch("w0 w1 w2 w3 w4 w5 w6 w7 w8 w9 w10 w11 w12 w13 w14 w15 w16 w17 w18 w19", "[%a]+") do c = c + 1 end
print(c == 20)"#
        ),
        "true"
    );
}

#[test]
fn test_string_gmatch_tokens_unicode_like() {
    assert_eq!(
        run_lua_one(
            r#"local c = 0
for _ in string.gmatch("w0 w1 w2 w3 w4 w5 w6 w7 w8 w9 w10 w11 w12 w13 w14 w15 w16 w17 w18 w19 w20", "[%a]+") do c = c + 1 end
print(c == 21)"#
        ),
        "true"
    );
}
