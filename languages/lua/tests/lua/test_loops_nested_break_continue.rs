use super::helpers::run_lua_one;

#[test]
fn test_loops_nested_break_continue_outer_still_runs_after_inner_break() {
    assert_eq!(
        run_lua_one("local total = 0\nfor outer = 1, 3 do\n  local skip = false\n  for inner = 1, 3 do\n    if inner == 2 then skip = true; break end\n    total = total + inner\n  end\n  if not skip then total = total + 10 end\nend\nprint(total)"),
        "24",
    );
}

#[test]
fn test_loops_nested_break_continue_outer_break_ends_immediately() {
    assert_eq!(
        run_lua_one("local count = 0\nfor outer = 1, 5 do\n  for inner = 1, 5 do\n    if outer == 2 and inner == 2 then break end\n    count = count + 1\n  end\n  if outer == 2 then break end\nend\nprint(count)"),
        "7",
    );
}

#[test]
fn test_loops_nested_break_continue_inner_break_only() {
    assert_eq!(
        run_lua_one("local total = 0\nfor outer = 1, 2 do\n  for inner = 1, 4 do\n    if inner == 3 then break end\n    total = total + outer\n  end\nend\nprint(total)"),
        "4",
    );
}

#[test]
fn test_loops_nested_break_continue_outer_gating() {
    assert_eq!(
        run_lua_one("local total = 0\nfor outer = 1, 5 do\n  local blocked = false\n  for inner = 1, 5 do\n    if inner == 4 then blocked = true; break end\n  end\n  if not blocked then total = total + outer end\nend\nprint(total)"),
        "15",
    );
}

#[test]
fn test_loops_nested_break_continue_inner_break_allows_next_outer() {
    assert_eq!(
        run_lua_one("local out = 0\nfor outer = 1, 4 do\n  for inner = 1, 6 do\n    if inner == 2 then break end\n    out = out + 1\n  end\nend\nprint(out)"),
        "4",
    );
}

#[test]
fn test_loops_nested_break_continue_skip_workaround() {
    assert_eq!(
        run_lua_one("local total = 0\nfor outer = 1, 3 do\n  for inner = 1, 4 do\n    if inner == 3 then total = total + 100; do end\n  end\n  total = total + outer\nend\nprint(total)"),
        "106",
    );
}

#[test]
fn test_loops_nested_break_continue_nested_break_only_inner() {
    assert_eq!(
        run_lua_one("local total = 0\nfor outer = 1, 2 do\n  for inner = 1, 4 do\n    if outer == 1 and inner == 2 then break end\n    total = total + inner\n  end\nend\nprint(total)"),
        "13",
    );
}

#[test]
fn test_loops_nested_break_continue_outer_after_inner_guard() {
    assert_eq!(
        run_lua_one("local sum = 0\nfor outer = 1, 3 do\n  local hit = false\n  for inner = 1, 3 do\n    if inner == 2 then hit = true end\n  end\n  if hit then sum = sum + 1 end\nend\nprint(sum)"),
        "3",
    );
}

#[test]
fn test_loops_nested_break_continue_double_nested() {
    assert_eq!(
        run_lua_one("local total = 0\nfor a = 1, 2 do\n  for b = 1, 2 do\n    for c = 1, 2 do\n      if a == 2 and b == 2 and c == 1 then break end\n      total = total + 1\n    end\n  end\nend\nprint(total)"),
        "7",
    );
}

#[test]
fn test_loops_nested_break_continue_inner_loop_counts_with_condition() {
    assert_eq!(
        run_lua_one("local total = 0\nfor outer = 1, 3 do\n  for inner = 1, 5 do\n    if inner == outer then break end\n    total = total + inner\n  end\nend\nprint(total)"),
        "12",
    );
}

#[test]
fn test_loops_nested_break_continue_break_on_first_match() {
    assert_eq!(
        run_lua_one("local total = 0\nfor outer = 1, 4 do\n  for inner = 1, 4 do\n    if inner == 1 then break end\n    total = total + 1\n  end\nend\nprint(total)"),
        "0",
    );
}

#[test]
fn test_loops_nested_break_continue_while_in_for() {
    assert_eq!(
        run_lua_one("local total = 0\nfor outer = 1, 3 do\n  local i = 0\n  while i < 4 do\n    i = i + 1\n    if i == 2 then break end\n    total = total + 1\n  end\nend\nprint(total)"),
        "3",
    );
}

#[test]
fn test_loops_nested_break_continue_outer_bounded_by_inner() {
    assert_eq!(
        run_lua_one("local total = 0\nfor outer = 1, 5 do\n  local hit = false\n  for inner = 1, 5 do\n    if inner == 4 then hit = true; break end\n  end\n  if not hit then total = total + 1 end\nend\nprint(total)"),
        "0",
    );
}

#[test]
fn test_loops_nested_break_continue_deep_loop_state() {
    assert_eq!(
        run_lua_one("local total = 0\nfor outer = 1, 3 do\n  for inner = 1, 3 do\n    for k = 1, 3 do\n      if outer == 2 and inner == 2 and k == 2 then total = total + 100; break end\n      total = total + 1\n    end\n  end\nend\nprint(total)"),
        "109",
    );
}

#[test]
fn test_loops_nested_break_continue_no_inner_break() {
    assert_eq!(
        run_lua_one("local count = 0\nfor outer = 1, 2 do\n  for inner = 1, 2 do\n    count = count + 1\n  end\nend\nprint(count)"),
        "4",
    );
}

#[test]
fn test_loops_nested_break_continue_outer_break_in_false_case() {
    assert_eq!(
        run_lua_one("local total = 0\nfor outer = 1, 4 do\n  for inner = 1, 4 do\n    if outer == 3 then break end\n    total = total + inner\n  end\n  if outer == 3 then total = total + 50 end\nend\nprint(total)"),
        "87",
    );
}

#[test]
fn test_loops_nested_break_continue_mix_tables() {
    assert_eq!(
        run_lua_one("local total = 0\nfor outer = 1, 2 do\n  local t = {1,2,3}\n  for _, v in ipairs(t) do\n    if v == 2 then break end\n    total = total + v\n  end\nend\nprint(total)"),
        "2",
    );
}

#[test]
fn test_loops_nested_break_continue_while_and_for() {
    assert_eq!(
        run_lua_one("local total = 0\nfor outer = 1, 3 do\n  local n = 0\n  while n < 4 do\n    n = n + 1\n    if n == 3 then break end\n    if outer == 2 then total = total + 10 end\n    total = total + 1\n  end\nend\nprint(total)"),
        "15",
    );
}

#[test]
fn test_loops_nested_break_continue_inner_and_outer_break_conditions() {
    assert_eq!(
        run_lua_one("local total = 0\nfor outer = 1, 6 do\n  for inner = 1, 6 do\n    if outer == 4 then break end\n    if inner == 3 then break end\n    total = total + 1\n  end\n  if outer == 4 then break end\nend\nprint(total)"),
        "8",
    );
}

#[test]
fn test_loops_nested_break_continue_zero_skipped_in_inner() {
    assert_eq!(
        run_lua_one("local total = 0\nfor outer = 1, 4 do\n  local n = 0\n  repeat\n    n = n + 1\n    if n == 1 then total = total + 1 else total = total + 2 end\n  until n == 2\n  if outer == 2 then total = total + 1 end\nend\nprint(total)"),
        "9",
    );
}

