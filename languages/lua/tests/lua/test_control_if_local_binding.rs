use super::helpers::run_lua_one;

#[test]
fn test_control_if_local_binding_true_branch_prints_value() {
    assert_eq!(run_lua_one("if true then local value = 4; print(value) end"), "4");
}

#[test]
fn test_control_if_local_binding_false_branch_not_bound() {
    assert_eq!(
        run_lua_one("if false then local value = 12 end print(value == nil and \"nil\" or \"bound\")"),
        "nil",
    );
}

#[test]
fn test_control_if_local_binding_elseif_keeps_that_branch() {
    assert_eq!(
        run_lua_one("if false then local value = 1 elseif true then local value = 3 print(value) end"),
        "3",
    );
}

#[test]
fn test_control_if_local_binding_nested_shadowing_prefers_inner() {
    assert_eq!(
        run_lua_one("if true then local a = 5; if true then local a = 2 print(a) end end"),
        "2",
    );
}

#[test]
fn test_control_if_local_binding_outer_block_stays_outside_shadow() {
    assert_eq!(
        run_lua_one("if true then local a = 5; do local a = 7 end print(a) end"),
        "5",
    );
}

#[test]
fn test_control_if_local_binding_local_function_invocation() {
    assert_eq!(
        run_lua_one("if true then local compute = function() return 11 end print(compute()) end"),
        "11",
    );
}

#[test]
fn test_control_if_local_binding_true_branch_with_local_mutation() {
    assert_eq!(
        run_lua_one("if true then local counter = 1; counter = counter + 2; print(counter) end"),
        "3",
    );
}

#[test]
fn test_control_if_local_binding_table_isolated_in_branch() {
    assert_eq!(
        run_lua_one("if true then local t = {a = 2, b = 3}; print(t.a + t.b) end"),
        "5",
    );
}

#[test]
fn test_control_if_local_binding_false_branch_has_its_own_local() {
    assert_eq!(
        run_lua_one("if false then local message = \"x\" else local message = \"ok\" print(message) end"),
        "ok",
    );
}

#[test]
fn test_control_if_local_binding_string_concat_in_local_branch() {
    assert_eq!(
        run_lua_one("if true then local p = \"a\" local q = \"b\" print(p .. q) end"),
        "ab",
    );
}

#[test]
fn test_control_if_local_binding_local_scope_used_in_nested_do() {
    assert_eq!(
        run_lua_one("if true then local total = 0; do local add = 5; total = total + add end print(total) end"),
        "5",
    );
}

#[test]
fn test_control_if_local_binding_local_boolean() {
    assert_eq!(run_lua_one("if true then local enabled = true; print(enabled) end"), "true");
}

#[test]
fn test_control_if_local_binding_elseif_else_selects_else_local() {
    assert_eq!(
        run_lua_one("if false then local x = 1 elseif false then local x = 2 else local x = 8 print(x) end"),
        "8",
    );
}

#[test]
fn test_control_if_local_binding_local_and_function_call_result() {
    assert_eq!(
        run_lua_one("if true then local f = function(v) return v * 2 end print(f(4)) end"),
        "8",
    );
}

#[test]
fn test_control_if_local_binding_local_kept_when_while_skips() {
    assert_eq!(
        run_lua_one("if true then local n = 9; while false do n = n + 1 end print(n) end"),
        "9",
    );
}

#[test]
fn test_control_if_local_binding_block_shadow_does_not_escape() {
    assert_eq!(
        run_lua_one(
            "if true then local x = 1; if true then local x = 2 end print(x == 1 and 1 or 0) end",
        ),
        "1",
    );
}

#[test]
fn test_control_if_local_binding_nested_if_inner_local_changes_count() {
    assert_eq!(
        run_lua_one("if true then local x = 1; if true then local y = 2 else local y = 3 end print(x + y) end"),
        "3",
    );
}

#[test]
fn test_control_if_local_binding_local_table_update() {
    assert_eq!(
        run_lua_one("if true then local t = {} t[\"a\"] = 1 t[\"a\"] = t[\"a\"] + 4 print(t[\"a\"]) end"),
        "5",
    );
}

#[test]
fn test_control_if_local_binding_local_returned_from_if_body() {
    assert_eq!(
        run_lua_one("if true then local n = 0; if true then n = 7 end print(n) end"),
        "7",
    );
}

#[test]
fn test_control_if_local_binding_local_branch_selects_default_when_nil() {
    assert_eq!(
        run_lua_one("if false then local value = 1 else local value = nil print(value == nil and \"nil\" or \"v\") end"),
        "nil",
    );
}

#[test]
fn test_control_if_local_binding_local_with_or_default() {
    assert_eq!(
        run_lua_one("if true then local value = nil; local out = value or 19 print(out) end"),
        "19",
    );
}

#[test]
fn test_control_if_local_binding_inner_local_arithmetic_chain() {
    assert_eq!(
        run_lua_one("if true then local base = 2; do local base = 3 local bonus = 4 print(base * bonus) end end"),
        "12",
    );
}

#[test]
fn test_control_if_local_binding_local_for_loop_within_branch() {
    assert_eq!(
        run_lua_one("if true then local total = 0; for i = 1, 3 do total = total + i end print(total) end"),
        "6",
    );
}

