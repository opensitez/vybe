use super::helpers::run_lua_one;

#[test]
fn test_loops_for_negative_step_basic_decrement() {
    assert_eq!(
        run_lua_one("local sum = 0\nfor i = 10, 1, -1 do sum = sum + i end\nprint(sum)"),
        "55",
    );
}

#[test]
fn test_loops_for_negative_step_every_other() {
    assert_eq!(
        run_lua_one("local sum = 0\nfor i = 10, 1, -2 do sum = sum + i end\nprint(sum)"),
        "30",
    );
}

#[test]
fn test_loops_for_negative_step_start_equals_end() {
    assert_eq!(
        run_lua_one("local count = 0\nfor i = 5, 5, -1 do count = count + 1 end\nprint(count)"),
        "1",
    );
}

#[test]
fn test_loops_for_negative_step_start_below_end_skips() {
    assert_eq!(
        run_lua_one("local count = 0\nfor i = 1, 5, -1 do count = count + 1 end\nprint(count)"),
        "0",
    );
}

#[test]
fn test_loops_for_negative_step_large_stride() {
    assert_eq!(
        run_lua_one("local sum = 0\nfor i = 25, 1, -7 do sum = sum + i end\nprint(sum)"),
        "58",
    );
}

#[test]
fn test_loops_for_negative_step_zero_stride_guarded_by_bound() {
    assert_eq!(
        run_lua_one("local count = 0\nfor i = 1, 1, 0 do count = count + 1 end\nprint(count)"),
        "1",
    );
}

#[test]
fn test_loops_for_negative_step_break_stops_early() {
    assert_eq!(
        run_lua_one("local sum = 0\nfor i = 9, 1, -1 do if i == 4 then break end sum = sum + i end\nprint(sum)"),
        "35",
    );
}

#[test]
fn test_loops_for_negative_step_nested_if() {
    assert_eq!(
        run_lua_one("local even = 0\nfor i = 12, 2, -2 do if i % 4 == 0 then even = even + 1 end end\nprint(even)"),
        "3",
    );
}

#[test]
fn test_loops_for_negative_step_fractional_step() {
    assert_eq!(
        run_lua_one("local sum = 0\nfor i = 10, 4, -3 do sum = sum + i end\nprint(sum)"),
        "21",
    );
}

#[test]
fn test_loops_for_negative_step_negative_start() {
    assert_eq!(
        run_lua_one("local sum = 0\nfor i = -2, -8, -2 do sum = sum + i end\nprint(sum)"),
        "-20",
    );
}

#[test]
fn test_loops_for_negative_step_step_three() {
    assert_eq!(
        run_lua_one("local sum = 0\nfor i = 20, 7, -3 do sum = sum + i end\nprint(sum)"),
        "70",
    );
}

#[test]
fn test_loops_for_negative_step_multiple_of_six() {
    assert_eq!(
        run_lua_one("local sum = 0\nfor i = 18, 0, -6 do sum = sum + i end\nprint(sum)"),
        "36",
    );
}

#[test]
fn test_loops_for_negative_step_count_items() {
    assert_eq!(
        run_lua_one("local count = 0\nfor i = 9, 2, -2 do count = count + 1 end\nprint(count)"),
        "4",
    );
}

#[test]
fn test_loops_for_negative_step_accumulate_with_odd_even() {
    assert_eq!(
        run_lua_one("local even = 0\nfor i = 11, 3, -2 do if i % 2 == 0 then even = even + i end end\nprint(even)"),
        "0",
    );
}

#[test]
fn test_loops_for_negative_step_sum_until_too_small() {
    assert_eq!(
        run_lua_one("local sum = 0\nfor i = 6, 0, -2 do if i < 3 then break end sum = sum + i end\nprint(sum)"),
        "10",
    );
}

#[test]
fn test_loops_for_negative_step_decrement_with_multiplier() {
    assert_eq!(
        run_lua_one("local out = 1\nfor i = 8, 2, -2 do out = out * i end\nprint(out)"),
        "384",
    );
}

#[test]
fn test_loops_for_negative_step_with_local_variable_step() {
    assert_eq!(
        run_lua_one("local sum = 0\nlocal step = -4\nfor i = 16, 1, step do sum = sum + i end\nprint(sum)"),
        "40",
    );
}

#[test]
fn test_loops_for_negative_step_then_if_adds_offsets() {
    assert_eq!(
        run_lua_one("local sum = 0\nfor i = 7, 0, -1 do if i > 4 then sum = sum + 1 end end\nprint(sum)"),
        "3",
    );
}

#[test]
fn test_loops_for_negative_step_large_range() {
    assert_eq!(
        run_lua_one("local count = 0\nfor i = 50, 41, -3 do count = count + 1 end\nprint(count)"),
        "4",
    );
}

#[test]
fn test_loops_for_negative_step_string_counter() {
    assert_eq!(
        run_lua_one("local out = \"\"\nfor i = 5, 1, -1 do out = out .. tostring(i) end\nprint(out)"),
        "54321",
    );
}

#[test]
fn test_loops_for_negative_step_nested_blocks() {
    assert_eq!(
        run_lua_one("local sum = 0\nfor i = 9, 1, -1 do do sum = sum + i end end\nprint(sum)"),
        "45",
    );
}

#[test]
fn test_loops_for_negative_step_uses_negative_step_with_continue_like_if() {
    assert_eq!(
        run_lua_one("local sum = 0\nfor i = 10, 1, -2 do if i == 6 then sum = sum + 0 else sum = sum + i end end\nprint(sum)"),
        "24",
    );
}
