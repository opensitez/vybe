lua_print! {
    test_repeat_basic => {
        "local i=1; local s=''; repeat s=s..i; i=i+1 until i>3; print(s)",
        "123"
    },
    test_repeat_executes_at_least_once => {
        "local i=10; local s=''; repeat s=s..i until true; print(s)",
        "10"
    },
    test_repeat_scope_of_until => {
        "local i=1; local s=''; repeat local j=i; s=s..j; i=i+1 until j>=3; print(s)",
        "123"
    },
    test_repeat_break => {
        "local i=1; local s=''; repeat s=s..i; if i==2 then break end; i=i+1 until false; print(s)",
        "12"
    },
    test_repeat_nested => {
        "local i=1; local s=''; repeat local j=1; repeat s=s..i..j; j=j+1 until j>2; i=i+1 until i>2; print(s)",
        "11122122"
    },
    test_repeat_closure_in_until => {
        "local i=1; repeat local j=i; local f=function() return j>2 end; i=i+1 until f(); print(i)",
        "4"
    },
    test_repeat_truthiness => {
        "local s=''; local t={1, 2, nil}; local i=1; repeat s=s..t[i]; i=i+1 until not t[i]; print(s)",
        "12"
    },
    repeat_local_in_body_visible_in_until_condition => {
        "local n = 0\nrepeat\n  local limit = 3\n  n = n + 1\nuntil n >= limit\nprint(n)\n",
        "3"
    },
    repeat_break_from_nested_loop_does_not_exit_outer => {
        "local i = 0\nrepeat\n  i = i + 1\n  for j = 1, 3 do\n    if j == 2 then break end\n  end\nuntil i == 3\nprint(i)\n",
        "3"
    },
    repeat_upvalue_captured_per_iteration => {
        "local out = ''\nlocal i = 0\nrepeat\n  i = i + 1\n  local x = i * 10\n  out = out .. x .. ','\nuntil i == 3\nprint(out)\n",
        "10,20,30,"
    },
    repeat_condition_reads_table_mutated_in_body => {
        "local t = {done = false}\nlocal n = 0\nrepeat\n  n = n + 1\n  if n == 4 then t.done = true end\nuntil t.done\nprint(n)\n",
        "4"
    },
    repeat_goto_skips_rest_of_body => {
        "local sum = 0\nlocal i = 0\nrepeat\n  i = i + 1\n  if i % 2 ~= 0 then sum = sum + i end\nuntil i == 6\nprint(sum)\n",
        "9"
    },
    repeat_executes_body_once_when_condition_immediately_true => {
        "local executed = false\nrepeat\n  executed = true\nuntil true\nprint(executed)\n",
        "true"
    },
    repeat_with_function_returning_condition => {
        "local count = 0\nlocal function should_stop()\n  count = count + 1\n  return count >= 5\nend\nrepeat\nuntil should_stop()\nprint(count)\n",
        "5"
    },
    repeat_nested_outer_break_exits_correctly => {
        "local result = 0\nlocal i = 0\nrepeat\n  i = i + 1\n  repeat\n    result = result + 1\n    break\n  until true\nuntil i == 3\nprint(result)\n",
        "3"
    } }
