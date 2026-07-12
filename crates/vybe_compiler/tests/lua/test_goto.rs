//! `goto` and labels — Lua 5.2+ manual §3.3.8.

lua_print! {
    goto_skips_unreachable_assignment => {
        "local x = 1\n goto finish\n x = 2\n ::finish::\n print(x)\n",
        "1"
    },
    goto_can_form_simple_loop => {
        "local n = 0\n::loop::\n n = n + 1\n if n < 3 then goto loop end\n print(n)\n",
        "3"
    },
    goto_forward_to_label => {
        "local s = \"\"\n goto two\n s = s .. \"1\"\n ::two::\n s = s .. \"2\"\n print(s)\n",
        "2"
    },
    break_and_goto_do_not_share_labels => {
        "local n = 0\nwhile n < 2 do\n  n = n + 1\n  goto done\nend\nn = 9\n::done::\nprint(n)\n",
        "1"
    },
    goto_across_local_declaration_scope => {
        "local x = 1\ngoto target\ndo\n  local y = 2\nend\n::target::\nprint(x)\n",
        "1"
    },
    goto_backward_reexecutes_local_initialization => {
        "local count = 0\n::start::\nlocal x = 10\ncount = count + 1\nx = x + count\nif count < 3 then goto start end\nprint(x)\n",
        "13"
    },
    goto_simulate_nested_break => {
        "local s = \"\"\nfor i = 1, 3 do\n  for j = 1, 3 do\n    if i * j == 4 then goto exit_all end\n    s = s .. i .. j .. \" \"\n  end\nend\n::exit_all::\nprint(s)\n",
        "11 12 13 21 "
    },
    goto_simulate_continue => {
        "local s = \"\"\nfor i = 1, 4 do\n  if i == 3 then goto skip end\n  s = s .. i\n  ::skip::\nend\nprint(s)\n",
        "124"
    },
    goto_nested_do_blocks_forward_jump => {
        "local x = 0\ndo\n  do\n    goto dest\n  end\nend\nx = 100\n::dest::\nprint(x)\n",
        "0"
    },
    goto_in_while_to_skip_iteration => {
        "local sum = 0\nlocal i = 0\nwhile i < 6 do\n  i = i + 1\n  if i % 2 == 0 then goto next end\n  sum = sum + i\n  ::next::\nend\nprint(sum)\n",
        "9"
    },
    goto_in_repeat_until_skips_body_rest => {
        "local n = 0\nlocal count = 0\nrepeat\n  n = n + 1\n  if n % 3 == 0 then goto skip end\n  count = count + 1\n  ::skip::\nuntil n == 6\nprint(count)\n",
        "4"
    },
    multiple_labels_in_same_block_both_reachable => {
        "local x = 0\ngoto first\n::first::\nx = 1\ngoto second\n::second::\nx = x + 10\nprint(x)\n",
        "11"
    },
    goto_over_global_function_definition_is_valid => {
        "goto after\nfunction skip_fn() return 99 end\n::after::\nprint(type(skip_fn))\n",
        "function"
    },
    goto_exits_nested_ipairs_iteration => {
        "local found = nil\nfor _, row in ipairs({{1,2},{3,4},{5,6}}) do\n  for _, v in ipairs(row) do\n    if v == 4 then found = v; goto done end\n  end\nend\n::done::\nprint(found)\n",
        "4"
    },
    goto_jumps_from_inside_if_to_after_end => {
        "local x = 0\nif true then\n  x = 1\n  goto after\n  x = 2\nend\n::after::\nprint(x)\n",
        "1"
    },
    goto_backward_accumulates_sum => {
        "local total = 0\nlocal i = 1\n::again::\ntotal = total + i\ni = i + 1\nif i <= 4 then goto again end\nprint(total)\n",
        "10"
    },
}
