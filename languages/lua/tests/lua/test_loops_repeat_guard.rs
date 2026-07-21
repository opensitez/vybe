use super::helpers::run_lua_one;

#[test]
fn test_loops_repeat_guard_executes_at_least_once() {
    assert_eq!(
        run_lua_one("local n = 0\nrepeat n = n + 1 until n > 0\nprint(n)"),
        "1",
    );
}

#[test]
fn test_loops_repeat_guard_runs_until_condition() {
    assert_eq!(
        run_lua_one("local n = 0\nrepeat n = n + 1 until n == 3\nprint(n)"),
        "3",
    );
}

#[test]
fn test_loops_repeat_guard_stops_on_false_guard() {
    assert_eq!(
        run_lua_one("local n = 0\nrepeat n = n + 2; if n > 7 then break end print(n) until false\nprint(n)"),
        "2",
    );
}

#[test]
fn test_loops_repeat_guard_with_conditional_mutation() {
    assert_eq!(
        run_lua_one("local value = 1\nrepeat\n  value = value * 2\nuntil value > 10\nprint(value)"),
        "16",
    );
}

#[test]
fn test_loops_repeat_guard_with_break() {
    assert_eq!(
        run_lua_one("local n = 0\nrepeat n = n + 1\nif n == 2 then break end\nuntil n > 10\nprint(n)"),
        "2",
    );
}

#[test]
fn test_loops_repeat_guard_string_concat_guard() {
    assert_eq!(
        run_lua_one("local s = \"\"\nrepeat s = s .. 'x' until #s > 2\nprint(s)"),
        "xxx",
    );
}

#[test]
fn test_loops_repeat_guard_false_condition_true_after_iterations() {
    assert_eq!(
        run_lua_one("local sum = 0\nrepeat sum = sum + 2; local next = sum > 4; until next\nprint(sum)"),
        "6",
    );
}

#[test]
fn test_loops_repeat_guard_nested_if() {
    assert_eq!(
        run_lua_one("local n = 0\nlocal t = 0\nrepeat n = n + 1; if n % 2 == 0 then t = t + 1 end until n == 4\nprint(t)"),
        "2",
    );
}

#[test]
fn test_loops_repeat_guard_boolean_guard_variable() {
    assert_eq!(
        run_lua_one("local n = 0\nlocal done = false\nrepeat\n  n = n + 1\n  if n >= 3 then done = true end\nuntil done\nprint(n)"),
        "3",
    );
}

#[test]
fn test_loops_repeat_guard_with_function() {
    assert_eq!(
        run_lua_one("local n = 0\nlocal function done(v) return v > 2 end\nrepeat n = n + 1 until done(n)\nprint(n)"),
        "3",
    );
}

#[test]
fn test_loops_repeat_guard_product_until_limit() {
    assert_eq!(
        run_lua_one("local n = 1\nrepeat n = n * 2; if n > 20 then break end until false\nprint(n)"),
        "32",
    );
}

#[test]
fn test_loops_repeat_guard_negative_step() {
    assert_eq!(
        run_lua_one("local n = 5\nrepeat n = n - 2 until n <= 1\nprint(n)"),
        "1",
    );
}

#[test]
fn test_loops_repeat_guard_countdown() {
    assert_eq!(
        run_lua_one("local n = 3\nlocal out = 0\nrepeat out = out + 1; n = n - 1 until n == 0\nprint(out)"),
        "3",
    );
}

#[test]
fn test_loops_repeat_guard_guard_on_table() {
    assert_eq!(
        run_lua_one("local t = {}\nrepeat table.insert(t, #t + 1) until #t == 2\nprint(#t)"),
        "2",
    );
}

#[test]
fn test_loops_repeat_guard_guard_with_nil_result() {
    assert_eq!(
        run_lua_one("local done = 0\nrepeat done = done + 1; local v = nil until v == nil\nprint(done)"),
        "1",
    );
}

#[test]
fn test_loops_repeat_guard_boolean_math() {
    assert_eq!(
        run_lua_one("local n = 0\nrepeat n = n + 1 until n * n > 10\nprint(n)"),
        "4",
    );
}

#[test]
fn test_loops_repeat_guard_conditional_string() {
    assert_eq!(
        run_lua_one("local out = ''\nrepeat out = out .. 'a' until #out > 1\nprint(out)"),
        "aa",
    );
}

#[test]
fn test_loops_repeat_guard_nested_repeat() {
    assert_eq!(
        run_lua_one("local n = 0\nlocal total = 0\nrepeat\n  n = n + 1\n  local inner = 0\n  repeat\n    inner = inner + 1\n    total = total + inner\n  until inner > 1\nuntil n > 2\nprint(total)"),
        "9",
    );
}

#[test]
fn test_loops_repeat_guard_zero_condition_immediate() {
    assert_eq!(
        run_lua_one("local n = 0\nrepeat n = n + 1 until true\nprint(n)"),
        "1",
    );
}

#[test]
fn test_loops_repeat_guard_with_table_length_condition() {
    assert_eq!(
        run_lua_one("local t = {}\nrepeat table.insert(t, 1) until #t == 1\nprint(#t)"),
        "1",
    );
}

#[test]
fn test_loops_repeat_guard_with_function_call_condition() {
    assert_eq!(
        run_lua_one("local n = 0\nlocal function done(v) return v >= 2 end\nrepeat n = n + 1 until done(n)\nprint(n)"),
        "2",
    );
}

#[test]
fn test_loops_repeat_guard_with_continue_like_guard() {
    assert_eq!(
        run_lua_one("local n = 0\nrepeat n = n + 1; if n == 2 then n = n + 1 else n = n end until n > 3\nprint(n)"),
        "4",
    );
}
