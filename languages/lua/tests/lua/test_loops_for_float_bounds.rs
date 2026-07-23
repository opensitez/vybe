use super::helpers::run_lua_one;

#[test]
fn test_loops_for_float_bounds_whole_number_representations() {
    assert_eq!(
        run_lua_one("local sum = 0\nfor i = 1.0, 7.0, 2.0 do sum = sum + i end\nprint(sum)"),
        "16",
    );
}

#[test]
fn test_loops_for_float_bounds_negative_direction() {
    assert_eq!(
        run_lua_one("local sum = 0\nfor i = 7.0, 1.0, -2.0 do sum = sum + i end\nprint(sum)"),
        "16",
    );
}

#[test]
fn test_loops_for_float_bounds_float_start_only_one_step() {
    assert_eq!(
        run_lua_one("local sum = 0\nfor i = 2.0, 4.0, 1.0 do sum = sum + i end\nprint(sum)"),
        "9",
    );
}

#[test]
fn test_loops_for_float_bounds_float_step_with_one_step() {
    assert_eq!(
        run_lua_one("local sum = 0\nfor i = 1.0, 3.0, 1.0 do sum = sum + i end\nprint(sum)"),
        "6",
    );
}

#[test]
fn test_loops_for_float_bounds_small_range() {
    assert_eq!(
        run_lua_one("local sum = 0\nfor i = 1.0, 2.0, 0.5 do sum = sum + 1 end\nprint(sum)"),
        "3",
    );
}

#[test]
fn test_loops_for_float_bounds_negative_fraction_range() {
    assert_eq!(
        run_lua_one("local sum = 0\nfor i = -1.0, -3.0, -1.0 do sum = sum + i end\nprint(sum)"),
        "-6",
    );
}

#[test]
fn test_loops_for_float_bounds_zero_based_float_bounds() {
    assert_eq!(
        run_lua_one("local sum = 0\nfor i = 0.0, 4.0, 2.0 do sum = sum + i end\nprint(sum)"),
        "6",
    );
}

#[test]
fn test_loops_for_float_bounds_float_boundaries_with_large_stride() {
    assert_eq!(
        run_lua_one(
            "local values = 0\nfor i = 10.0, 4.0, -3.0 do values = values + 1 end\nprint(values)"
        ),
        "3",
    );
}

#[test]
fn test_loops_for_float_bounds_single_float_iteration() {
    assert_eq!(
        run_lua_one(
            "local count = 0\nfor i = 2.5, 2.5, 1.0 do count = count + 1 end\nprint(count)"
        ),
        "1",
    );
}

#[test]
fn test_loops_for_float_bounds_float_limit_non_integer_step() {
    assert_eq!(
        run_lua_one("local sum = 0\nfor i = 1.0, 3.0, 1.0 do sum = sum + i end\nprint(sum)"),
        "6",
    );
}

#[test]
fn test_loops_for_float_bounds_accumulate_string_index() {
    assert_eq!(
        run_lua_one(
            "local out = \"\"\nfor i = 1.0, 5.0, 2.0 do out = out .. tostring(i) end\nprint(out)"
        ),
        "135",
    );
}

#[test]
fn test_loops_for_float_bounds_stringifying_float_step() {
    assert_eq!(
        run_lua_one(
            "local out = \"\"\nfor i = 1.0, 3.0, 1.0 do out = out .. tostring(i) .. ';' end\nprint(out)"
        ),
        "1;2;3;",
    );
}

#[test]
fn test_loops_for_float_bounds_decimal_offset_addition() {
    assert_eq!(
        run_lua_one("local sum = 0\nfor i = 1.2, 3.6, 1.2 do sum = sum + 1 end\nprint(sum)"),
        "3",
    );
}

#[test]
fn test_loops_for_float_bounds_decimal_offset_count() {
    assert_eq!(
        run_lua_one(
            "local count = 0\nfor i = 0.0, 1.0, 0.25 do count = count + 1 end\nprint(count)"
        ),
        "5",
    );
}

#[test]
fn test_loops_for_float_bounds_decimal_stride_skips_fractionals() {
    assert_eq!(
        run_lua_one(
            "local count = 0\nfor i = 0.0, 2.0, 0.5 do if i > 0 then count = count + 1 end end\nprint(count)"
        ),
        "4",
    );
}

#[test]
fn test_loops_for_float_bounds_float_breaks_midway() {
    assert_eq!(
        run_lua_one(
            "local sum = 0\nfor i = 5.0, 1.0, -1.0 do if i == 3 then break end sum = sum + i end\nprint(sum)"
        ),
        "9",
    );
}

#[test]
fn test_loops_for_float_bounds_local_step_variable() {
    assert_eq!(
        run_lua_one(
            "local sum = 0\nlocal step = 2.0\nfor i = 1.0, 6.0, step do sum = sum + i end\nprint(sum)"
        ),
        "9",
    );
}

#[test]
fn test_loops_for_float_bounds_fractional_limit() {
    assert_eq!(
        run_lua_one("local sum = 0\nfor i = 1.0, 2.5, 0.75 do sum = sum + 1 end\nprint(sum)"),
        "3",
    );
}

#[test]
fn test_loops_for_float_bounds_negative_bound_inclusive() {
    assert_eq!(
        run_lua_one(
            "local values = 0\nfor i = 0.0, -4.0, -2.0 do values = values + 1 end\nprint(values)"
        ),
        "3",
    );
}

#[test]
fn test_loops_for_float_bounds_float_nested_in_if() {
    assert_eq!(
        run_lua_one(
            "local sum = 0\nif true then for i = 4.0, 2.0, -1.0 do sum = sum + i end end\nprint(sum)"
        ),
        "9",
    );
}

#[test]
fn test_loops_for_float_bounds_float_default_step_behavior() {
    assert_eq!(
        run_lua_one("local count = 0\nfor i = 1.0, 3.0 do count = count + 1 end\nprint(count)"),
        "3",
    );
}

#[test]
fn test_loops_for_float_bounds_float_large_bounds() {
    assert_eq!(
        run_lua_one("local sum = 0\nfor i = 1.0, 10.0, 3.0 do sum = sum + i end\nprint(sum)"),
        "22",
    );
}
