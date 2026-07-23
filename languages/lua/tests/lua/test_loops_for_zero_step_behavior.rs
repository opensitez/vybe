use super::helpers::run_lua_one;

#[test]
fn test_loops_for_zero_step_behavior_no_iteration_when_lower_bound_greater() {
    assert_eq!(
        run_lua_one("local count = 0\nfor i = 1, 0, 0 do count = count + 1 end\nprint(count)"),
        "0",
    );
}

#[test]
fn test_loops_for_zero_step_behavior_no_iteration_when_negative_range() {
    assert_eq!(
        run_lua_one("local count = 0\nfor i = 0, -1, 0 do count = count + 1 end\nprint(count)"),
        "0",
    );
}

#[test]
fn test_loops_for_zero_step_behavior_single_step_omitted() {
    assert_eq!(
        run_lua_one("local sum = 0\nfor i = 1, 4 do sum = sum + i end\nprint(sum)"),
        "10",
    );
}

#[test]
fn test_loops_for_zero_step_behavior_single_step_by_one() {
    assert_eq!(
        run_lua_one("local count = 0\nfor i = 3, 3 do count = count + 1 end\nprint(count)"),
        "1",
    );
}

#[test]
fn test_loops_for_zero_step_behavior_default_step_guarded_by_negative_bound() {
    assert_eq!(
        run_lua_one("local count = 0\nfor i = 5, 1 do count = count + 1 end\nprint(count)"),
        "0",
    );
}

#[test]
fn test_loops_for_zero_step_behavior_default_step_reverse_bound() {
    assert_eq!(
        run_lua_one("local count = 0\nfor i = 1, 5 do count = count + 1 end\nprint(count)"),
        "5",
    );
}

#[test]
fn test_loops_for_zero_step_behavior_zero_step_with_hoisted_guard() {
    assert_eq!(
        run_lua_one(
            "local value = 0\nlocal start = 0\nlocal stop = 1\nif stop < start then for i = start, stop, 0 do value = value + 1 end end\nprint(value)"
        ),
        "0",
    );
}

#[test]
fn test_loops_for_zero_step_behavior_zero_step_false_branch() {
    assert_eq!(
        run_lua_one(
            "local total = 0\nif false then for i = 1, 2, 0 do total = total + 1 end else total = 4 end\nprint(total)"
        ),
        "4",
    );
}

#[test]
fn test_loops_for_zero_step_behavior_zero_step_positive_like_guard() {
    assert_eq!(
        run_lua_one(
            "local total = 0\nfor i = 2, 1, 0 do total = total + 1 end\nif total == 0 then print(1) else print(0) end"
        ),
        "1",
    );
}

#[test]
fn test_loops_for_zero_step_behavior_zero_step_never_entering_loop() {
    assert_eq!(
        run_lua_one("local total = 0\nfor i = 2, 2, -1 do total = total + 1 end\nprint(total)"),
        "1",
    );
}

#[test]
fn test_loops_for_zero_step_behavior_zero_step_nested_no_loop() {
    assert_eq!(
        run_lua_one(
            "local total = 0\nfor i = 1, 0, 0 do total = total + 1 end\nfor j = 1, 3 do total = total + 1 end\nprint(total)"
        ),
        "3",
    );
}

#[test]
fn test_loops_for_zero_step_behavior_zero_step_terminates_due_else() {
    assert_eq!(
        run_lua_one(
            "local flag = false\nfor i = 1, 0, 0 do flag = true end\nprint(flag == true and \"yes\" or \"no\")"
        ),
        "no",
    );
}

#[test]
fn test_loops_for_zero_step_behavior_local_bounds_areolated() {
    assert_eq!(
        run_lua_one(
            "local start = 9\nlocal stop = 3\nlocal sum = 0\nif start > stop then for i = start, stop, 0 do sum = sum + 1 end end\nprint(sum)"
        ),
        "0",
    );
}

#[test]
fn test_loops_for_zero_step_behavior_local_bounds_equal_no_stride() {
    assert_eq!(
        run_lua_one(
            "local start = 5\nlocal stop = 5\nlocal count = 0\nif false then for i = start, stop, 0 do count = count + 1 end end\nprint(count)"
        ),
        "0",
    );
}

#[test]
fn test_loops_for_zero_step_behavior_conditional_zero_step() {
    assert_eq!(
        run_lua_one(
            "local count = 0\nif true then for i = 4, 2, 0 do count = count + 1 end end\nprint(count == 0 and \"empty\" or \"nonempty\")"
        ),
        "empty",
    );
}

#[test]
fn test_loops_for_zero_step_behavior_default_step_range_with_gap() {
    assert_eq!(
        run_lua_one("local sum = 0\nfor i = 1, 3 do sum = sum + i end\nprint(sum)"),
        "6",
    );
}

#[test]
fn test_loops_for_zero_step_behavior_default_step_full_span() {
    assert_eq!(
        run_lua_one("local out = \"\"\nfor i = 3, 6 do out = out .. tostring(i) end\nprint(out)"),
        "3456",
    );
}

#[test]
fn test_loops_for_zero_step_behavior_step_expression_zero_guarded() {
    assert_eq!(
        run_lua_one(
            "local count = 0\nlocal step = 0\nlocal run_zero = false\nif run_zero then for i = 3, 1, step do count = count + 1 end else for i = 3, 1, -1 do count = count + 1 end end\nprint(count)"
        ),
        "3",
    );
}

#[test]
fn test_loops_for_zero_step_behavior_non_zero_equivalent() {
    assert_eq!(
        run_lua_one(
            "local total = 0\nlocal step = -1\nfor i = 10, 8, step do total = total + i end\nprint(total)"
        ),
        "27",
    );
}

#[test]
fn test_loops_for_zero_step_behavior_zero_step_when_limit_too_low() {
    assert_eq!(
        run_lua_one("local value = 0\nfor i = 2, 0, 0 do value = value + i end\nprint(value)"),
        "0",
    );
}

#[test]
fn test_loops_for_zero_step_behavior_nested_scope_skip() {
    assert_eq!(
        run_lua_one(
            "local total = 0\nif false then for i = 3, 1, 0 do total = total + 1 end else do total = total + 2 end end\nprint(total)"
        ),
        "2",
    );
}

#[test]
fn test_loops_for_zero_step_behavior_zero_step_range_guarded_by_if_true() {
    assert_eq!(
        run_lua_one(
            "local value = \"\"\nif 1 > 2 then for i = 1, 0, 0 do value = value .. \"x\" end end\nprint(value == \"\" and \"empty\" or \"filled\")"
        ),
        "empty",
    );
}
