use super::helpers::run_lua_one;

#[test]
fn test_loops_for_generic_pairs_mut_base_sum() {
    assert_eq!(
        run_lua_one("local total = 0\nlocal source = {a = 1, b = 2, c = 3}\nfor _, value in pairs(source) do total = total + value end\nprint(total)"),
        "6",
    );
}

#[test]
fn test_loops_for_generic_pairs_mut_mutate_existing_values() {
    assert_eq!(
        run_lua_one("local t = {a = 1, b = 2}\nfor k, v in pairs(t) do t[k] = v + 1 end\nprint(t.a + t.b)"),
        "5",
    );
}

#[test]
fn test_loops_for_generic_pairs_mut_update_nested_tables() {
    assert_eq!(
        run_lua_one("local t = {a = {x = 1}, b = {x = 2}}\nfor _, v in pairs(t) do v.x = v.x + 1 end\nprint(t.a.x + t.b.x)"),
        "6",
    );
}

#[test]
fn test_loops_for_generic_pairs_mut_count_keys() {
    assert_eq!(
        run_lua_one("local count = 0\nlocal t = {a = 1, b = 2, c = 3}\nfor _ in pairs(t) do count = count + 1 end\nprint(count)"),
        "3",
    );
}

#[test]
fn test_loops_for_generic_pairs_mut_sum_of_lengthy_keys() {
    assert_eq!(
        run_lua_one("local t = {left = 4, right = 8}\nlocal sum = 0\nfor key, value in pairs(t) do if type(key) == \"string\" then sum = sum + value end end\nprint(sum)"),
        "12",
    );
}

#[test]
fn test_loops_for_generic_pairs_mut_mutate_numeric_key_stringified() {
    assert_eq!(
        run_lua_one("local t = {['1'] = 10, ['2'] = 20}\nfor key, value in pairs(t) do t[key] = value * 2 end\nprint((t['1'] + t['2']) / 2)"),
        "30",
    );
}

#[test]
fn test_loops_for_generic_pairs_mut_boolean_values() {
    assert_eq!(
        run_lua_one("local t = {a = true, b = false, c = true}\nlocal sum = 0\nfor _, value in pairs(t) do if value then sum = sum + 1 end end\nprint(sum)"),
        "2",
    );
}

#[test]
fn test_loops_for_generic_pairs_mut_merge_string_values() {
    assert_eq!(
        run_lua_one("local t = {a = 'x', b = 'y'}\nlocal out = ''\nfor key, value in pairs(t) do if #out == 0 then out = key .. '=' .. value else out = out end end\nprint((out == 'a=x' or out == 'b=y') and 1 or 0)"),
        "1",
    );
}

#[test]
fn test_loops_for_generic_pairs_mut_boolean_gate_on_key() {
    assert_eq!(
        run_lua_one("local t = {one = 1, two = 2, three = 3}\nlocal hits = 0\nfor key, value in pairs(t) do if key == 'two' then hits = hits + value end end\nprint(hits)"),
        "2",
    );
}

#[test]
fn test_loops_for_generic_pairs_mut_sum_after_mutation() {
    assert_eq!(
        run_lua_one("local t = {a = 1, b = 2, c = 3}\nfor _, value in pairs(t) do value = value + 5 end\nprint(t.a + t.b + t.c)"),
        "6",
    );
}

#[test]
fn test_loops_for_generic_pairs_mut_double_number_values() {
    assert_eq!(
        run_lua_one("local t = {a = 2, b = 4}\nfor key, value in pairs(t) do t[key] = value * 2 end\nprint(t.a + t.b)"),
        "12",
    );
}

#[test]
fn test_loops_for_generic_pairs_mut_nested_loop_sum() {
    assert_eq!(
        run_lua_one("local matrix = {x = {1,2}, y = {3,4}}\nlocal out = 0\nfor _, row in pairs(matrix) do for _, value in ipairs(row) do out = out + value end end\nprint(out)"),
        "10",
    );
}

#[test]
fn test_loops_for_generic_pairs_mut_rebind_existing_function() {
    assert_eq!(
        run_lua_one("local t = {f = function(a) return a + 1 end}\nfor key, value in pairs(t) do if key == 'f' then t[key] = function(a) return value(a) * 2 end end\nprint(t.f(2))"),
        "6",
    );
}

#[test]
fn test_loops_for_generic_pairs_mut_boolean_value_sums() {
    assert_eq!(
        run_lua_one("local t = {a = true, b = true, c = false}\nlocal total = 0\nfor _, value in pairs(t) do if value then total = total + 1 else total = total - 1 end end\nprint(total)"),
        "1",
    );
}

#[test]
fn test_loops_for_generic_pairs_mut_mutate_only_targets() {
    assert_eq!(
        run_lua_one("local t = {a = 1, b = 2}\nfor key, value in pairs(t) do if key == 'a' then t[key] = value + 10 end end\nprint(t.a + t.b)"),
        "13",
    );
}

#[test]
fn test_loops_for_generic_pairs_mut_zero_values() {
    assert_eq!(
        run_lua_one("local t = {a = 0, b = -1, c = 1}\nlocal count = 0\nfor _, value in pairs(t) do if value == 0 then count = count + 1 end end\nprint(count)"),
        "1",
    );
}

#[test]
fn test_loops_for_generic_pairs_mut_total_string_lengths() {
    assert_eq!(
        run_lua_one("local t = {a = 'hi', b = 'x', c = 'lo'}\nlocal total = 0\nfor _, value in pairs(t) do if type(value) == 'string' then total = total + #value end end\nprint(total)"),
        "4",
    );
}

#[test]
fn test_loops_for_generic_pairs_mut_pairs_after_setting_defaults() {
    assert_eq!(
        run_lua_one("local t = {x = 10}\nfor key in pairs(t) do t[key .. '_done'] = true end\nprint((t.x_done == nil) and 0 or 1)"),
        "0",
    );
}

#[test]
fn test_loops_for_generic_pairs_mut_iteration_does_not_change_key_count() {
    assert_eq!(
        run_lua_one("local t = {a = 1, b = 2}\nlocal before = 0\nfor _ in pairs(t) do before = before + 1 end\nfor k, v in pairs(t) do t[k] = v + 1 end\nlocal after = 0\nfor _ in pairs(t) do after = after + 1 end\nprint(before == after and 1 or 0)"),
        "1",
    );
}

#[test]
fn test_loops_for_generic_pairs_mut_sum_with_mutation_inside_loop() {
    assert_eq!(
        run_lua_one("local t = {a = 1, b = 2}\nlocal total = 0\nfor k, v in pairs(t) do t[k] = v * 2; total = total + t[k] end\nprint(total)"),
        "6",
    );
}

#[test]
fn test_loops_for_generic_pairs_mut_false_guard_stops_usage() {
    assert_eq!(
        run_lua_one("local t = {a = 1, b = 2, c = 3}\nif false then for k,v in pairs(t) do t[k]=v+1 end end\nprint(t.a + t.b + t.c)"),
        "6",
    );
}

#[test]
fn test_loops_for_generic_pairs_mut_key_projection() {
    assert_eq!(
        run_lua_one("local t = {alpha = 1, beta = 2}\nlocal out = ''\nfor k, value in pairs(t) do out = out .. k end\nprint(#out == 9 and 1 or 0)"),
        "1",
    );
}

