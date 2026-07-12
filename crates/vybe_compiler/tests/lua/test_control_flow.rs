use super::helpers::{parse_ok, run_lua_one};

#[test]
fn if_then_prints_branch() {
    let out = run_lua_one("if 1 < 2 then print(\"yes\") end\n");
    assert_eq!(out, "yes");
}

#[test]
fn if_else_picks_else() {
    let out = run_lua_one("if false then print(\"a\") else print(\"b\") end\n");
    assert_eq!(out, "b");
}

#[test]
fn elseif_chain() {
    let src = "if false then print(1)\nelseif true then print(2)\nelse print(3) end\n";
    let out = run_lua_one(src);
    assert_eq!(out, "2");
}

#[test]
fn while_loop_counts() {
    let src = "local i = 0\nwhile i < 3 do\n  i = i + 1\nend\nprint(i)\n";
    let out = run_lua_one(src);
    assert_eq!(out, "3");
}

#[test]
fn repeat_until_runs_once() {
    let src = "local n = 0\nrepeat\n  n = n + 1\nuntil n >= 1\nprint(n)\n";
    let out = run_lua_one(src);
    assert_eq!(out, "1");
}

#[test]
fn break_inside_while() {
    let src =
        "local i = 0\nwhile true do\n  i = i + 1\n  if i == 2 then break end\nend\nprint(i)\n";
    let out = run_lua_one(src);
    assert_eq!(out, "2");
}

#[test]
fn return_from_function() {
    let src = "function f()\n  return 7\nend\nprint(f())\n";
    let out = run_lua_one(src);
    assert_eq!(out, "7");
}

#[test]
fn nested_if_parses() {
    parse_ok("if true then if false then print(1) end end\n");
}

// ── Spec gaps: truthiness, elseif/else, loops (Lua 5.x manual §3.3) ────────

lua_print! {
    if_nil_takes_else => {
        "if nil then print(\"a\") else print(\"b\") end\n",
        "b"
    },
    if_zero_is_truthy => {
        "if 0 then print(\"yes\") else print(\"no\") end\n",
        "yes"
    },
    elseif_all_false_hits_else => {
        "if false then print(1) elseif false then print(2) else print(3) end\n",
        "3"
    },
    if_false_leaves_local_unchanged => {
        "local x = 1\nif false then x = 2 end\nprint(x)\n",
        "1"
    },
    nested_if_inner_else_runs => {
        "if true then if false then print(1) else print(2) end end\n",
        "2"
    },
    while_zero_never_runs_body => {
        "local n = 0\nwhile 0 do n = n + 1 end\nprint(n)\n",
        "0"
    },
    repeat_until_true_exits_after_one_body => {
        "local n = 0\nrepeat n = n + 1 until true\nprint(n)\n",
        "1"
    },
    nested_while_accumulates => {
        "local sum = 0\nlocal i = 1\nwhile i <= 3 do\n  local j = 1\n  while j <= 2 do\n    sum = sum + 1\n    j = j + 1\n  end\n  i = i + 1\nend\nprint(sum)\n",
        "6"
    },
    repeat_until_exits_when_condition_becomes_true => {
        "local n = 0\nrepeat\n  n = n + 1\nuntil n >= 3\nprint(n)\n",
        "3"
    },
    numeric_for_loop_counts_up => {
        "local sum = 0\nfor i = 1, 5 do\n  sum = sum + i\nend\nprint(sum)\n",
        "15"
    },
    if_without_else_skips_when_false => {
        "local x = 1\nif false then x = 2 end\nprint(x)\n",
        "1"
    },
    elseif_skips_else_when_matched => {
        "if false then print(1)\nelseif true then print(2)\nelse print(3) end\n",
        "2"
    },
    repeat_body_always_runs_once_before_condition => {
        "local n = 0\nrepeat n = n + 1 until n >= 1\nprint(n)\n",
        "1"
    },
    break_inside_repeat_exits_loop => {
        "local n = 0\nrepeat\n  n = n + 1\n  if n == 2 then break end\nuntil n >= 9\nprint(n)\n",
        "2"
    },
    while_with_empty_string_condition_runs_until_break => {
        "local c = 0\nwhile \"\" do\n  c = c + 1\n  if c == 2 then break end\nend\nprint(c)\n",
        "2"
    },
    nested_if_without_else_on_inner => {
        "if true then if false then print(1) end print(2) end\n",
        "2"
    },
    if_true_picks_then_over_else => {
        "if true then print(\"t\") else print(\"f\") end\n",
        "t"
    },
    if_uses_local_in_condition => {
        "local n = 3\nif n > 2 then print(\"big\") end\n",
        "big"
    },
    while_repeats_until_condition_false => {
        "local i = 0\nwhile i < 2 do i = i + 1 end\nprint(i)\n",
        "2"
    },
    repeat_always_runs_once_minimum => {
        "local ran = false\nrepeat ran = true until true\nprint(ran)\n",
        "true"
    },
    break_inside_while_stops_loop => {
        "local i = 0\nwhile i < 9 do\n  i = i + 1\n  if i == 3 then break end\nend\nprint(i)\n",
        "3"
    },
    elseif_first_match_wins => {
        "if false then print(1)\nelseif true then print(2)\nelseif true then print(3) end\n",
        "2"
    },
    if_nested_in_while_body => {
        "local n = 0\nwhile n < 3 do\n  n = n + 1\n  if n == 2 then print(\"hit\") end\nend\n",
        "hit"
    },
    return_exits_function_early => {
        "function f()\n  if true then return 5 end\n  return 0\nend\nprint(f())\n",
        "5"
    },
    local_updated_inside_while => {
        "local sum = 0\nlocal i = 1\nwhile i <= 2 do\n  sum = sum + i\n  i = i + 1\nend\nprint(sum)\n",
        "3"
    },
    if_elseif_else_with_local_result => {
        "local n = 2\nlocal label\nif n < 0 then label = \"neg\"\nelseif n == 0 then label = \"zero\"\nelse label = \"pos\" end\nprint(label)\n",
        "pos"
    },
    while_with_break_on_condition => {
        "local i = 0\nwhile true do\n  i = i + 1\n  if i >= 3 then break end\nend\nprint(i)\n",
        "3"
    },
    repeat_until_with_local_counter => {
        "local tries = 0\nrepeat tries = tries + 1 until tries == 3\nprint(tries)\n",
        "3"
    },
    for_loop_prints_last_value => {
        "local last = 0\nfor i = 1, 4 do last = i end\nprint(last)\n",
        "4"
    },
    nested_if_selects_inner_branch => {
        "if true then if true then print(\"inner\") else print(\"outer\") end end\n",
        "inner"
    },
    if_not_nil_check_common_idiom => {
        "local v = 1\nif v ~= nil then print(\"set\") end\n",
        "set"
    },
    elseif_short_circuit_after_match => {
        "local x = 2\nif x == 1 then print(1)\nelseif x == 2 then print(2)\nelse print(3) end\n",
        "2"
    },
    if_zero_condition_is_truthy => {
        "if 0 then print(\"yes\") else print(\"no\") end\n",
        "yes"
    },
    if_empty_string_condition_is_truthy => {
        "if \"\" then print(\"yes\") else print(\"no\") end\n",
        "yes"
    },
    elseif_chain_picks_first_true_predicate => {
        "local n = 5\nif n < 0 then print(\"neg\")\nelseif n == 0 then print(\"zero\")\nelseif n > 0 then print(\"pos\") end\n",
        "pos"
    },
    else_runs_when_all_branches_false => {
        "if false then print(1) elseif false then print(2) else print(3) end\n",
        "3"
    },
    logical_and_in_if_condition => {
        "if true and true then print(\"ok\") end\n",
        "ok"
    },
    logical_or_in_if_condition => {
        "if false or true then print(\"ok\") end\n",
        "ok"
    },
    not_operator_negates_false => {
        "if not false then print(\"ok\") end\n",
        "ok"
    },
    nested_if_else_resolves_inner_first => {
        "if true then if false then print(1) else print(2) end else print(3) end\n",
        "2"
    },
    if_with_comparison_on_strings => {
        "if \"a\" < \"b\" then print(\"ordered\") end\n",
        "ordered"
    },
    if_with_nil_explicit_check => {
        "local v = nil\nif v == nil then print(\"unset\") end\n",
        "unset"
    },
    if_with_false_explicit_check => {
        "local flag = false\nif flag == false then print(\"off\") end\n",
        "off"
    },
    guard_clause_style_early_exit => {
        "local function f(x)\n  if x == nil then return \"nil\" end\n  return \"ok\"\nend\nprint(f(nil))\n",
        "nil"
    },
    elseif_chain_falls_to_else_when_all_fail => {
        "local x = 5\nif x == 1 then print('a')\nelseif x == 2 then print('b')\nelseif x == 3 then print('c')\nelse print('none')\nend\n",
        "none"
    },
    if_condition_calls_function_for_truthiness => {
        "local calls = 0\nlocal function check() calls = calls + 1; return calls > 2 end\nwhile not check() do end\nprint(calls)\n",
        "3"
    },
    nested_if_inside_else_branch => {
        "local x = 5\nlocal result\nif x < 3 then\n  result = 'low'\nelse\n  if x < 7 then\n    result = 'mid'\n  else\n    result = 'high'\n  end\nend\nprint(result)\n",
        "mid"
    },
    break_from_for_loop_preserves_outer_var => {
        "local outer = 'preserved'\nfor i = 1, 5 do\n  if i == 3 then break end\nend\nprint(outer)\n",
        "preserved"
    },
    break_only_exits_innermost_loop => {
        "local sum = 0\nfor i = 1, 3 do\n  for j = 1, 3 do\n    if j == 2 then break end\n    sum = sum + j\n  end\nend\nprint(sum)\n",
        "3"
    },
    while_with_or_condition_evaluates_second_part => {
        "local a, b = false, true\nlocal ran = false\nwhile a or b do\n  ran = true\n  b = false\nend\nprint(ran)\n",
        "true"
    },
    return_value_from_inside_if_branch => {
        "local function classify(n)\n  if n < 0 then return 'negative'\n  elseif n == 0 then return 'zero'\n  else return 'positive'\n  end\nend\nprint(classify(-5) .. ',' .. classify(0) .. ',' .. classify(3))\n",
        "negative,zero,positive"
    },
}
