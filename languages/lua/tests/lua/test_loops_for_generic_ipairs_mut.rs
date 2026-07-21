use super::helpers::run_lua_one;

#[test]
fn test_loops_for_generic_ipairs_mut_basic_sum() {
    assert_eq!(
        run_lua_one("local sum = 0\nfor i, value in ipairs({1,2,3}) do sum = sum + value end\nprint(sum)"),
        "6",
    );
}

#[test]
fn test_loops_for_generic_ipairs_mut_mutate_in_place() {
    assert_eq!(
        run_lua_one("local t = {1,2,3}\nfor i, value in ipairs(t) do t[i] = value * 2 end\nprint(t[1] + t[2] + t[3])"),
        "12",
    );
}

#[test]
fn test_loops_for_generic_ipairs_mut_count_indices() {
    assert_eq!(
        run_lua_one("local count = 0\nfor i in ipairs({7,8,9}) do count = count + 1 end\nprint(count)"),
        "3",
    );
}

#[test]
fn test_loops_for_generic_ipairs_mut_stop_at_first_hole() {
    assert_eq!(
        run_lua_one("local t = {1,2,3}\nt[2] = nil\nlocal count = 0\nfor _, value in ipairs(t) do if value then count = count + 1 end end\nprint(count)"),
        "1",
    );
}

#[test]
fn test_loops_for_generic_ipairs_mut_index_sum() {
    assert_eq!(
        run_lua_one("local sum = 0\nfor i, value in ipairs({4,5,6}) do if i == 2 then sum = sum + value end end\nprint(sum)"),
        "5",
    );
}

#[test]
fn test_loops_for_generic_ipairs_mut_even_filter() {
    assert_eq!(
        run_lua_one("local sum = 0\nfor _, value in ipairs({1,2,3,4,5}) do if value % 2 == 0 then sum = sum + value end end\nprint(sum)"),
        "6",
    );
}

#[test]
fn test_loops_for_generic_ipairs_mut_reverse() {
    assert_eq!(
        run_lua_one("local t = {1,2,3}\nlocal out = 0\nfor i, value in ipairs(t) do out = out + value end\nt = {3,2,1}\nprint((t[1] + out) == 7 and 1 or 0)"),
        "0",
    );
}

#[test]
fn test_loops_for_generic_ipairs_mut_build_suffix() {
    assert_eq!(
        run_lua_one("local out = ''\nfor i, value in ipairs({1,2}) do out = out .. tostring(i) .. ':' .. tostring(value) .. ';' end\nprint(out)"),
        "1:1;2:2;",
    );
}

#[test]
fn test_loops_for_generic_ipairs_mut_mutate_indexed_table() {
    assert_eq!(
        run_lua_one("local t = {1,1,1}\nfor i, value in ipairs(t) do t[i] = value + i end\nprint(t[1] + t[2] + t[3])"),
        "9",
    );
}

#[test]
fn test_loops_for_generic_ipairs_mut_nested_loop_pairs() {
    assert_eq!(
        run_lua_one("local total = 0\nfor _, row in ipairs({{1,2},{3,4}}) do for _, value in ipairs(row) do total = total + value end end\nprint(total)"),
        "10",
    );
}

#[test]
fn test_loops_for_generic_ipairs_mut_string_table() {
    assert_eq!(
        run_lua_one("local t = {'a','b','c'}\nlocal out = ''\nfor _, v in ipairs(t) do out = out .. v end\nprint(out)"),
        "abc",
    );
}

#[test]
fn test_loops_for_generic_ipairs_mut_boolean_list() {
    assert_eq!(
        run_lua_one("local t = {true, false, true}\nlocal false_count = 0\nfor _, value in ipairs(t) do if value == false then false_count = false_count + 1 end end\nprint(false_count)"),
        "1",
    );
}

#[test]
fn test_loops_for_generic_ipairs_mut_count_with_zero() {
    assert_eq!(
        run_lua_one("local t = {0,0,1}\nlocal sum = 0\nfor _, value in ipairs(t) do sum = sum + 1 end\nprint(sum)"),
        "3",
    );
}

#[test]
fn test_loops_for_generic_ipairs_mut_skip_by_condition() {
    assert_eq!(
        run_lua_one("local sum = 0\nfor i, value in ipairs({1,2,3,4}) do if i > 2 then sum = sum + value end end\nprint(sum)"),
        "7",
    );
}

#[test]
fn test_loops_for_generic_ipairs_mut_assign_none() {
    assert_eq!(
        run_lua_one("local seen = 0\nfor _, value in ipairs({1, nil, 3}) do if value then seen = seen + 1 end end\nprint(seen)"),
        "1",
    );
}

#[test]
fn test_loops_for_generic_ipairs_mut_mutate_with_factor() {
    assert_eq!(
        run_lua_one("local t = {2,4,6}\nfor i, value in ipairs(t) do t[i] = value / 2 end\nprint(t[1] + t[2] + t[3])"),
        "6",
    );
}

#[test]
fn test_loops_for_generic_ipairs_mut_stringify_all() {
    assert_eq!(
        run_lua_one("local out = 0\nfor _, value in ipairs({1,2,3}) do out = out + value end\nprint(out)"),
        "6",
    );
}

#[test]
fn test_loops_for_generic_ipairs_mut_negative_numbers() {
    assert_eq!(
        run_lua_one("local sum = 0\nfor _, value in ipairs({-1,-2,-3}) do sum = sum + value end\nprint(sum)"),
        "-6",
    );
}

#[test]
fn test_loops_for_generic_ipairs_mut_mutable_tables() {
    assert_eq!(
        run_lua_one("local t = {{v=1},{v=2}}\nfor _, item in ipairs(t) do item.v = item.v + 1 end\nprint(t[1].v + t[2].v)"),
        "5",
    );
}

#[test]
fn test_loops_for_generic_ipairs_mut_sum_even_indexed() {
    assert_eq!(
        run_lua_one("local sum = 0\nfor i, value in ipairs({10,20,30,40}) do if i % 2 == 0 then sum = sum + value end end\nprint(sum)"),
        "60",
    );
}

#[test]
fn test_loops_for_generic_ipairs_mut_assign_accumulator() {
    assert_eq!(
        run_lua_one("local products = 1\nfor _, value in ipairs({1,2,3}) do products = products * value end\nprint(products)"),
        "6",
    );
}
