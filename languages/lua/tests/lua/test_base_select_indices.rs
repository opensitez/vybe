use super::helpers::run_lua_one;

#[test]
fn test_case_01() {
    assert_eq!(run_lua_one(r#"local idx = -3; local ok, _ = pcall(function() return select(idx, 10, 20, 30) end); local expect = (idx == 1 or idx == 2 or idx == 3 or idx == -1 or idx == -2 or idx == -3); print(ok == expect)"#), "true");
}

#[test]
fn test_case_02() {
    assert_eq!(run_lua_one(r#"local idx = -2; local ok, _ = pcall(function() return select(idx, 10, 20, 30) end); local expect = (idx == 1 or idx == 2 or idx == 3 or idx == -1 or idx == -2 or idx == -3); print(ok == expect)"#), "true");
}

#[test]
fn test_case_03() {
    assert_eq!(run_lua_one(r#"local idx = -1; local ok, _ = pcall(function() return select(idx, 10, 20, 30) end); local expect = (idx == 1 or idx == 2 or idx == 3 or idx == -1 or idx == -2 or idx == -3); print(ok == expect)"#), "true");
}

#[test]
fn test_case_04() {
    assert_eq!(run_lua_one(r#"local idx = 0; local ok, _ = pcall(function() return select(idx, 10, 20, 30) end); local expect = (idx == 1 or idx == 2 or idx == 3 or idx == -1 or idx == -2 or idx == -3); print(ok == expect)"#), "true");
}

#[test]
fn test_case_05() {
    assert_eq!(run_lua_one(r#"local idx = 1; local ok, _ = pcall(function() return select(idx, 10, 20, 30) end); local expect = (idx == 1 or idx == 2 or idx == 3 or idx == -1 or idx == -2 or idx == -3); print(ok == expect)"#), "true");
}

#[test]
fn test_case_06() {
    assert_eq!(run_lua_one(r#"local idx = 2; local ok, _ = pcall(function() return select(idx, 10, 20, 30) end); local expect = (idx == 1 or idx == 2 or idx == 3 or idx == -1 or idx == -2 or idx == -3); print(ok == expect)"#), "true");
}

#[test]
fn test_case_07() {
    assert_eq!(run_lua_one(r#"local idx = 3; local ok, _ = pcall(function() return select(idx, 10, 20, 30) end); local expect = (idx == 1 or idx == 2 or idx == 3 or idx == -1 or idx == -2 or idx == -3); print(ok == expect)"#), "true");
}

#[test]
fn test_case_08() {
    assert_eq!(run_lua_one(r#"local idx = 4; local ok, _ = pcall(function() return select(idx, 10, 20, 30) end); local expect = (idx == 1 or idx == 2 or idx == 3 or idx == -1 or idx == -2 or idx == -3); print(ok == expect)"#), "true");
}

#[test]
fn test_case_09() {
    assert_eq!(run_lua_one(r#"local idx = 5; local ok, _ = pcall(function() return select(idx, 10, 20, 30) end); local expect = (idx == 1 or idx == 2 or idx == 3 or idx == -1 or idx == -2 or idx == -3); print(ok == expect)"#), "true");
}

#[test]
fn test_case_10() {
    assert_eq!(run_lua_one(r#"local idx = -4; local ok, _ = pcall(function() return select(idx, 10, 20, 30) end); local expect = (idx == 1 or idx == 2 or idx == 3 or idx == -1 or idx == -2 or idx == -3); print(ok == expect)"#), "true");
}

#[test]
fn test_case_11() {
    assert_eq!(run_lua_one(r#"local idx = -5; local ok, _ = pcall(function() return select(idx, 10, 20, 30) end); local expect = (idx == 1 or idx == 2 or idx == 3 or idx == -1 or idx == -2 or idx == -3); print(ok == expect)"#), "true");
}

#[test]
fn test_case_12() {
    assert_eq!(run_lua_one(r#"local idx = 6; local ok, _ = pcall(function() return select(idx, 10, 20, 30) end); local expect = (idx == 1 or idx == 2 or idx == 3 or idx == -1 or idx == -2 or idx == -3); print(ok == expect)"#), "true");
}

#[test]
fn test_case_13() {
    assert_eq!(run_lua_one(r#"local idx = 7; local ok, _ = pcall(function() return select(idx, 10, 20, 30) end); local expect = (idx == 1 or idx == 2 or idx == 3 or idx == -1 or idx == -2 or idx == -3); print(ok == expect)"#), "true");
}

#[test]
fn test_case_14() {
    assert_eq!(run_lua_one(r#"local idx = 8; local ok, _ = pcall(function() return select(idx, 10, 20, 30) end); local expect = (idx == 1 or idx == 2 or idx == 3 or idx == -1 or idx == -2 or idx == -3); print(ok == expect)"#), "true");
}

#[test]
fn test_case_15() {
    assert_eq!(run_lua_one(r#"local idx = 9; local ok, _ = pcall(function() return select(idx, 10, 20, 30) end); local expect = (idx == 1 or idx == 2 or idx == 3 or idx == -1 or idx == -2 or idx == -3); print(ok == expect)"#), "true");
}

#[test]
fn test_case_16() {
    assert_eq!(run_lua_one(r#"local idx = 10; local ok, _ = pcall(function() return select(idx, 10, 20, 30) end); local expect = (idx == 1 or idx == 2 or idx == 3 or idx == -1 or idx == -2 or idx == -3); print(ok == expect)"#), "true");
}

#[test]
fn test_case_17() {
    assert_eq!(run_lua_one(r#"local idx = 11; local ok, _ = pcall(function() return select(idx, 10, 20, 30) end); local expect = (idx == 1 or idx == 2 or idx == 3 or idx == -1 or idx == -2 or idx == -3); print(ok == expect)"#), "true");
}

#[test]
fn test_case_18() {
    assert_eq!(run_lua_one(r#"local idx = 12; local ok, _ = pcall(function() return select(idx, 10, 20, 30) end); local expect = (idx == 1 or idx == 2 or idx == 3 or idx == -1 or idx == -2 or idx == -3); print(ok == expect)"#), "true");
}

#[test]
fn test_case_19() {
    assert_eq!(run_lua_one(r#"local idx = 13; local ok, _ = pcall(function() return select(idx, 10, 20, 30) end); local expect = (idx == 1 or idx == 2 or idx == 3 or idx == -1 or idx == -2 or idx == -3); print(ok == expect)"#), "true");
}

#[test]
fn test_case_20() {
    assert_eq!(run_lua_one(r#"local idx = 14; local ok, _ = pcall(function() return select(idx, 10, 20, 30) end); local expect = (idx == 1 or idx == 2 or idx == 3 or idx == -1 or idx == -2 or idx == -3); print(ok == expect)"#), "true");
}

