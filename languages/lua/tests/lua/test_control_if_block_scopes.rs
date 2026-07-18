use super::helpers::run_lua_one;

#[test]
fn test_control_if_block_scopes_simple_block_accumulates() {
    assert_eq!(
        run_lua_one("if true then local sum = 0; do local x = 1; sum = sum + x end print(sum) end"),
        "1",
    );
}

#[test]
fn test_control_if_block_scopes_nested_block_keeps_outer() {
    assert_eq!(
        run_lua_one("if true then local sum = 0; do local x = 2; sum = sum + x; do local x = 5 sum = sum + x end end print(sum) end"),
        "7",
    );
}

#[test]
fn test_control_if_block_scopes_block_local_does_not_escape() {
    assert_eq!(
        run_lua_one("if true then local x = 1; do local y = 10 end print(x) end"),
        "1",
    );
}

#[test]
fn test_control_if_block_scopes_nested_blocks_with_two_locals() {
    assert_eq!(
        run_lua_one("if true then do local left = 2 do local right = 3 print(left + right) end end"),
        "5",
    );
}

#[test]
fn test_control_if_block_scopes_else_and_blocks() {
    assert_eq!(
        run_lua_one("if false then print(0) else do local a = 4 do local b = 5 print(a + b) end end end"),
        "9",
    );
}

#[test]
fn test_control_if_block_scopes_repeated_blocks_sum() {
    assert_eq!(
        run_lua_one("if true then local sum = 0; do local v = 1 sum = sum + v end do local v = 2 sum = sum + v end print(sum) end"),
        "3",
    );
}

#[test]
fn test_control_if_block_scopes_shadowed_variable_in_inner_block() {
    assert_eq!(
        run_lua_one("if true then local label = 1 do local label = 10 end print(label) end"),
        "1",
    );
}

#[test]
fn test_control_if_block_scopes_multiple_block_scopes() {
    assert_eq!(
        run_lua_one("if true then local sum = 0; do local x = 1 sum = sum + x end do local x = 2 sum = sum + x end print(sum) end"),
        "3",
    );
}

#[test]
fn test_control_if_block_scopes_conditional_block_scope() {
    assert_eq!(
        run_lua_one("if true then local base = 3; if base > 1 then do local base = base + 7 print(base) end end end"),
        "10",
    );
}

#[test]
fn test_control_if_block_scopes_block_return_with_loop() {
    assert_eq!(
        run_lua_one("if true then local total = 0; do for i = 1, 3 do total = total + i end end print(total) end"),
        "6",
    );
}

#[test]
fn test_control_if_block_scopes_blocked_assignment_outside() {
    assert_eq!(
        run_lua_one("if true then local x = 1 do local x = 2 end print(x == 1 and 1 or 0) end"),
        "1",
    );
}

#[test]
fn test_control_if_block_scopes_block_table_building() {
    assert_eq!(
        run_lua_one("if true then local t = {} do t[\"a\"] = 4 t[\"b\"] = 6 end print(t[\"a\"] + t[\"b\"]) end"),
        "10",
    );
}

#[test]
fn test_control_if_block_scopes_block_scopes_mixed_types() {
    assert_eq!(
        run_lua_one("if true then local count = 0; do local s = \"x\"; if s == \"x\" then count = count + 1 end end print(count) end"),
        "1",
    );
}

#[test]
fn test_control_if_block_scopes_inner_block_uses_outer() {
    assert_eq!(
        run_lua_one("if true then local base = 2; do local offset = 3 print(base + offset) end end"),
        "5",
    );
}

#[test]
fn test_control_if_block_scopes_block_with_if_else() {
    assert_eq!(
        run_lua_one("if true then local total = 0 do if total == 0 then total = 4 else total = 0 end print(total) end end"),
        "4",
    );
}

#[test]
fn test_control_if_block_scopes_local_scopes_chain() {
    assert_eq!(
        run_lua_one("if true then local a = 1 do local a = 2 do local a = 3 print(a + 10) end print(a + 20) end print(a + 30) end"),
        "13",
    );
}

#[test]
fn test_control_if_block_scopes_function_calls_inside_block() {
    assert_eq!(
        run_lua_one("if true then local f = function(v) return v + 1 end do print(f(3)) end"),
        "4",
    );
}

#[test]
fn test_control_if_block_scopes_block_sum_of_strings() {
    assert_eq!(
        run_lua_one("if true then local out = \"\" do local part = \"hi\" out = out .. part end print(out) end"),
        "hi",
    );
}

#[test]
fn test_control_if_block_scopes_block_local_booleans() {
    assert_eq!(
        run_lua_one("if true then local ok = false do local seen = true print(ok == false and seen == true and \"true\" or \"false\") end"),
        "true",
    );
}

#[test]
fn test_control_if_block_scopes_block_scope_after_loop_iteration() {
    assert_eq!(
        run_lua_one("if true then local sum = 0 do for i = 1, 2 do sum = sum + i end end print(sum) end"),
        "3",
    );
}

#[test]
fn test_control_if_block_scopes_nested_blocks_short_circuit_style() {
    assert_eq!(
        run_lua_one("if true then local value = 1 do if true then do local value = 8; value = value + 1 end end print(value) end"),
        "1",
    );
}

#[test]
fn test_control_if_block_scopes_block_scope_with_repeat() {
    assert_eq!(
        run_lua_one("if true then local x = 0 do repeat x = x + 1 until x >= 1 end print(x) end"),
        "1",
    );
}
