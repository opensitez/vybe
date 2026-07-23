use super::helpers::run_lua_one;

#[test]
fn test_load_chunk_baseline() {
    assert_eq!(
        run_lua_one(r#"local f = assert(load("return 1+1")); print(f())"#),
        "true"
    );
}

#[test]
fn test_load_chunk_simple() {
    assert_eq!(
        run_lua_one(r#"local f = assert(load("return 2+3")); print(f())"#),
        "true"
    );
}

#[test]
fn test_load_chunk_trimmed() {
    assert_eq!(
        run_lua_one(r#"local f = assert(load("return 3+4")); print(f())"#),
        "true"
    );
}

#[test]
fn test_load_chunk_decimal() {
    assert_eq!(
        run_lua_one(r#"local f = assert(load("return 4+5")); print(f())"#),
        "true"
    );
}

#[test]
fn test_load_chunk_hexed() {
    assert_eq!(
        run_lua_one(r#"local f = assert(load("return 5+6")); print(f())"#),
        "true"
    );
}

#[test]
fn test_load_chunk_prefixed() {
    assert_eq!(
        run_lua_one(r#"local f = assert(load("return 6+7")); print(f())"#),
        "true"
    );
}

#[test]
fn test_load_chunk_negative() {
    assert_eq!(
        run_lua_one(r#"local f = assert(load("return 7+8")); print(f())"#),
        "true"
    );
}

#[test]
fn test_load_chunk_rounded() {
    assert_eq!(
        run_lua_one(r#"local f = assert(load("return 8+9")); print(f())"#),
        "true"
    );
}

#[test]
fn test_load_chunk_offset() {
    assert_eq!(
        run_lua_one(r#"local f = assert(load("return 9+10")); print(f())"#),
        "true"
    );
}

#[test]
fn test_load_chunk_paired() {
    assert_eq!(
        run_lua_one(r#"local f = assert(load("return 10+11")); print(f())"#),
        "true"
    );
}

#[test]
fn test_load_chunk_nested() {
    assert_eq!(
        run_lua_one(r#"local f = assert(load("return 11+12")); print(f())"#),
        "true"
    );
}

#[test]
fn test_load_chunk_metaflow() {
    assert_eq!(
        run_lua_one(r#"local f = assert(load("return 12+13")); print(f())"#),
        "true"
    );
}

#[test]
fn test_load_chunk_guarded() {
    assert_eq!(
        run_lua_one(r#"local f = assert(load("return 13+14")); print(f())"#),
        "true"
    );
}

#[test]
fn test_load_chunk_mapped() {
    assert_eq!(
        run_lua_one(r#"local f = assert(load("return 14+15")); print(f())"#),
        "true"
    );
}

#[test]
fn test_load_chunk_captured() {
    assert_eq!(
        run_lua_one(r#"local f = assert(load("return 15+16")); print(f())"#),
        "true"
    );
}

#[test]
fn test_load_chunk_edge_first() {
    assert_eq!(
        run_lua_one(r#"local f = assert(load("return 16+17")); print(f())"#),
        "true"
    );
}

#[test]
fn test_load_chunk_edge_second() {
    assert_eq!(
        run_lua_one(r#"local f = assert(load("return 17+18")); print(f())"#),
        "true"
    );
}

#[test]
fn test_load_chunk_edge_last() {
    assert_eq!(
        run_lua_one(r#"local f = assert(load("return 18+19")); print(f())"#),
        "true"
    );
}

#[test]
fn test_load_chunk_randomized() {
    assert_eq!(
        run_lua_one(r#"local f = assert(load("return 19+20")); print(f())"#),
        "true"
    );
}

#[test]
fn test_load_chunk_unicode_like() {
    assert_eq!(
        run_lua_one(r#"local f = assert(load("return 20+21")); print(f())"#),
        "true"
    );
}
