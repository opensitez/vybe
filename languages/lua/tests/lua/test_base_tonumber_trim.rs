use super::helpers::run_lua_one;

#[test]
fn test_tonumber_trim_baseline() {
    assert_eq!(
        run_lua_one(
            r#"print(tonumber("
	 0 
") == 0)"#
        ),
        "true"
    );
}

#[test]
fn test_tonumber_trim_simple() {
    assert_eq!(
        run_lua_one(
            r#"print(tonumber("
	 1 
") == 1)"#
        ),
        "true"
    );
}

#[test]
fn test_tonumber_trim_trimmed() {
    assert_eq!(
        run_lua_one(
            r#"print(tonumber("
	 -2 
") == -2)"#
        ),
        "true"
    );
}

#[test]
fn test_tonumber_trim_decimal() {
    assert_eq!(
        run_lua_one(
            r#"print(tonumber("
	 3.5 
") == 3.5)"#
        ),
        "true"
    );
}

#[test]
fn test_tonumber_trim_hexed() {
    assert_eq!(
        run_lua_one(
            r#"print(tonumber("
	 -4.25 
") == -4.25)"#
        ),
        "true"
    );
}

#[test]
fn test_tonumber_trim_prefixed() {
    assert_eq!(
        run_lua_one(
            r#"print(tonumber("
	 10 
") == 10)"#
        ),
        "true"
    );
}

#[test]
fn test_tonumber_trim_negative() {
    assert_eq!(
        run_lua_one(
            r#"print(tonumber("
	 -11 
") == -11)"#
        ),
        "true"
    );
}

#[test]
fn test_tonumber_trim_rounded() {
    assert_eq!(
        run_lua_one(
            r#"print(tonumber("
	 100 
") == 100)"#
        ),
        "true"
    );
}

#[test]
fn test_tonumber_trim_offset() {
    assert_eq!(
        run_lua_one(
            r#"print(tonumber("
	 -200 
") == -200)"#
        ),
        "true"
    );
}

#[test]
fn test_tonumber_trim_paired() {
    assert_eq!(
        run_lua_one(
            r#"print(tonumber("
	 255 
") == 255)"#
        ),
        "true"
    );
}

#[test]
fn test_tonumber_trim_nested() {
    assert_eq!(
        run_lua_one(
            r#"print(tonumber("
	 1.25 
") == 1.25)"#
        ),
        "true"
    );
}

#[test]
fn test_tonumber_trim_metaflow() {
    assert_eq!(
        run_lua_one(
            r#"print(tonumber("
	 7 
") == 7)"#
        ),
        "true"
    );
}

#[test]
fn test_tonumber_trim_guarded() {
    assert_eq!(
        run_lua_one(
            r#"print(tonumber("
	 8 
") == 8)"#
        ),
        "true"
    );
}

#[test]
fn test_tonumber_trim_mapped() {
    assert_eq!(
        run_lua_one(
            r#"print(tonumber("
	 9 
") == 9)"#
        ),
        "true"
    );
}

#[test]
fn test_tonumber_trim_captured() {
    assert_eq!(
        run_lua_one(
            r#"print(tonumber("
	 12.75 
") == 12.75)"#
        ),
        "true"
    );
}

#[test]
fn test_tonumber_trim_edge_first() {
    assert_eq!(
        run_lua_one(
            r#"print(tonumber("
	 3 
") == 3)"#
        ),
        "true"
    );
}

#[test]
fn test_tonumber_trim_edge_second() {
    assert_eq!(
        run_lua_one(
            r#"print(tonumber("
	 -3 
") == -3)"#
        ),
        "true"
    );
}

#[test]
fn test_tonumber_trim_edge_last() {
    assert_eq!(
        run_lua_one(
            r#"print(tonumber("
	 42 
") == 42)"#
        ),
        "true"
    );
}

#[test]
fn test_tonumber_trim_randomized() {
    assert_eq!(
        run_lua_one(
            r#"print(tonumber("
	 -99 
") == -99)"#
        ),
        "true"
    );
}

#[test]
fn test_tonumber_trim_unicode_like() {
    assert_eq!(
        run_lua_one(
            r#"print(tonumber("
	 12345 
") == 12345)"#
        ),
        "true"
    );
}
