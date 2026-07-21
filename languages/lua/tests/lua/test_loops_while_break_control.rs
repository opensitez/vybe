use super::helpers::run_lua_one;

#[test]
fn test_loops_while_break_control_break_immediately() {
    assert_eq!(
        run_lua_one("local sum = 0\nlocal i = 0\nwhile true do break end\nprint(sum)"),
        "0",
    );
}

#[test]
fn test_loops_while_break_control_break_after_count() {
    assert_eq!(
        run_lua_one("local sum = 0\nlocal i = 0\nwhile true do i = i + 1; if i > 3 then break end sum = sum + i end\nprint(sum)"),
        "6",
    );
}

#[test]
fn test_loops_while_break_control_break_with_condition() {
    assert_eq!(
        run_lua_one("local count = 0\nwhile count < 10 do count = count + 1; if count == 5 then break end end\nprint(count)"),
        "5",
    );
}

#[test]
fn test_loops_while_break_control_no_break_to_end() {
    assert_eq!(
        run_lua_one("local sum = 0\nlocal i = 1\nwhile i <= 4 do sum = sum + i; i = i + 1 end\nprint(sum)"),
        "10",
    );
}

#[test]
fn test_loops_while_break_control_nested_if_break() {
    assert_eq!(
        run_lua_one("local count = 0\nwhile true do count = count + 1; if count == 2 then if false then break end end if count == 4 then break end end\nprint(count)"),
        "4",
    );
}

#[test]
fn test_loops_while_break_control_break_skips_update() {
    assert_eq!(
        run_lua_one("local value = 0\nlocal i = 0\nwhile i < 5 do i = i + 1; if i == 3 then break end value = value + 10 end\nprint(i .. ':' .. value)"),
        "3:20",
    );
}

#[test]
fn test_loops_while_break_control_multiple_break_points() {
    assert_eq!(
        run_lua_one("local i = 0\nlocal total = 0\nwhile i < 20 do i = i + 2; if i == 4 then total = total + 1 elseif i == 10 then break end total = total + i end\nprint(total)"),
        "21",
    );
}

#[test]
fn test_loops_while_break_control_while_false_after_break_guard() {
    assert_eq!(
        run_lua_one("local active = true\nlocal count = 0\nwhile active do count = count + 1; if count > 2 then active = false end if count == 1 then break end end\nprint(count)"),
        "1",
    );
}

#[test]
fn test_loops_while_break_control_break_and_continue_style() {
    assert_eq!(
        run_lua_one("local i = 0\nlocal sum = 0\nwhile i < 6 do i = i + 1; if i == 4 then break end sum = sum + i end\nprint(sum)"),
        "6",
    );
}

#[test]
fn test_loops_while_break_control_inner_break_condition() {
    assert_eq!(
        run_lua_one("local sum = 0\nlocal i = 0\nwhile i < 100 do i = i + 1; if i % 3 == 0 then if i > 6 then break end end sum = sum + 1 end\nprint(sum)"),
        "8",
    );
}

#[test]
fn test_loops_while_break_control_break_on_zero_hit() {
    assert_eq!(
        run_lua_one("local i = 5\nwhile i > 0 do if i == 0 then break end i = i - 1 end\nprint(i)"),
        "0",
    );
}

#[test]
fn test_loops_while_break_control_break_and_reach_end() {
    assert_eq!(
        run_lua_one("local i = 0\nwhile i < 5 do i = i + 1; if i == 3 then break end end\nprint(i)"),
        "3",
    );
}

#[test]
fn test_loops_while_break_control_break_not_taken() {
    assert_eq!(
        run_lua_one("local i = 0\nwhile i < 2 do i = i + 1 end\nprint(i)"),
        "2",
    );
}

#[test]
fn test_loops_while_break_control_boolean_total() {
    assert_eq!(
        run_lua_one("local i = 0\nlocal ok = false\nwhile i < 4 do i = i + 1; if i == 2 then ok = true end if i == 3 then break end end\nprint(ok and \"true\" or \"false\")"),
        "true",
    );
}

#[test]
fn test_loops_while_break_control_break_after_nested_operation() {
    assert_eq!(
        run_lua_one("local i = 0\nlocal total = 0\nwhile true do i = i + 1; if i == 1 then total = total + 1 else total = total + 2 end if total > 5 then break end end\nprint(total)"),
        "7",
    );
}

#[test]
fn test_loops_while_break_control_nested_table_updates() {
    assert_eq!(
        run_lua_one("local t = {1,2,3}\nlocal i = 0\nwhile i < 5 do i = i + 1; table.insert(t, i); if #t > 5 then break end end\nprint(#t)"),
        "6",
    );
}

#[test]
fn test_loops_while_break_control_even_only_count() {
    assert_eq!(
        run_lua_one("local i = 0\nlocal even = 0\nwhile i < 10 do i = i + 1; if i % 2 == 0 then even = even + 1 end if i == 9 then break end end\nprint(even)"),
        "4",
    );
}

#[test]
fn test_loops_while_break_control_string_counter() {
    assert_eq!(
        run_lua_one("local i = 0\nlocal out = ''\nwhile true do i = i + 1; out = out .. tostring(i); if i >= 3 then break end end\nprint(out)"),
        "123",
    );
}

#[test]
fn test_loops_while_break_control_flag_toggle() {
    assert_eq!(
        run_lua_one("local run = true\nlocal i = 0\nwhile run do i = i + 1; if i == 1 then i = i + 1 end; if i > 3 then break end; run = false end\nprint(i)"),
        "2",
    );
}

#[test]
fn test_loops_while_break_control_math_progress() {
    assert_eq!(
        run_lua_one("local x = 1\nwhile x < 100 do x = x * 2; if x == 8 then break end end\nprint(x)"),
        "8",
    );
}

#[test]
fn test_loops_while_break_control_last_value() {
    assert_eq!(
        run_lua_one("local n = 0\nwhile n < 6 do n = n + 1; if n == 5 then n = 9 end end\nprint(n)"),
        "9",
    );
}

#[test]
fn test_loops_while_break_control_break_after_false_path() {
    assert_eq!(
        run_lua_one("local i = 0\nwhile i < 10 do if i == 4 then break end i = i + 1 end\nprint(i)"),
        "4",
    );
}
