//! Iteration — numeric `for`, `ipairs`, `pairs` (Lua 5.x manual §3.3.5, §3.3.6).

lua_print! {
    numeric_for_inclusive_end => {
        "local last=0\nfor i=3,5 do last=i end\nprint(last)\n",
        "5"
    },
    numeric_for_with_step_skips_values => {
        "local s=\"\"\nfor i=1,6,2 do s=s..i end\nprint(s)\n",
        "135"
    },
    numeric_for_negative_step_counts_down => {
        "local s=\"\"\nfor i=3,1,-1 do s=s..i end\nprint(s)\n",
        "321"
    },
    numeric_for_zero_step_is_invalid_at_runtime => {
        "local ok,err=pcall(function() for i=1,3,0 do end end)\nprint(ok)\n",
        "false"
    },
    break_exits_numeric_for_early => {
        "local sum=0\nfor i=1,100 do\n  if i>4 then break end\n  sum=sum+i\nend\nprint(sum)\n",
        "10"
    },
    ipairs_visits_array_part_in_order => {
        "local t={10,20,30}\nlocal s=0\nfor _,v in ipairs(t) do s=s+v end\nprint(s)\n",
        "60"
    },
    ipairs_stops_at_first_nil_gap => {
        "local t={}\nt[1]=1\nt[2]=nil\nt[3]=3\nlocal n=0\nfor _ in ipairs(t) do n=n+1 end\nprint(n)\n",
        "1"
    },
    pairs_visits_hash_keys => {
        "local t={x=1,y=2}\nlocal n=0\nfor _ in pairs(t) do n=n+1 end\nprint(n)\n",
        "2"
    },
    generic_for_with_manual_iterator => {
        "local function iter(_,i) i=i+1 if i>2 then return nil end return i,i end\nlocal s=0\nfor _,v in iter,nil,0 do s=s+v end\nprint(s)\n",
        "3"
    },
    numeric_for_start_greater_than_end_runs_zero_times => {
        "local n = 0\nfor i = 5, 1 do n = n + 1 end\nprint(n)\n",
        "0"
    },
    pairs_includes_sequence_and_hash_keys => {
        "local t = {1, a = 2}\nlocal count = 0\nfor _ in pairs(t) do count = count + 1 end\nprint(count)\n",
        "2"
    },
    ipairs_does_not_visit_hash_keys => {
        "local t = {x = 1, 2, 3}\nlocal sum = 0\nfor _, v in ipairs(t) do sum = sum + v end\nprint(sum)\n",
        "5"
    },
    numeric_for_loop_variable_is_local => {
        "for i = 1, 1 do local i = 9 end\nprint(i)\n",
        "nil"
    },
    ipairs_builds_comma_separated_list => {
        "local t = {\"a\", \"b\", \"c\"}\nlocal s = \"\"\nfor i, v in ipairs(t) do\n  s = s .. (i > 1 and \",\" or \"\") .. v\nend\nprint(s)\n",
        "a,b,c"
    },
    pairs_reads_key_value_from_hash => {
        "local t = {name = \"lua\"}\nfor k, v in pairs(t) do print(k .. \"=\" .. v) end\n",
        "name=lua"
    },
    numeric_for_step_two_every_other => {
        "local s = \"\"\nfor i = 2, 6, 2 do s = s .. i end\nprint(s)\n",
        "246"
    },
    for_loop_with_ipairs_to_find_max => {
        "local t = {3, 9, 1}\nlocal max = t[1]\nfor _, v in ipairs(t) do if v > max then max = v end end\nprint(max)\n",
        "9"
    },
    while_iterator_manual_index => {
        "local t = {4, 5, 6}\nlocal i, sum = 1, 0\nwhile t[i] do sum = sum + t[i] i = i + 1 end\nprint(sum)\n",
        "15"
    },
    numeric_for_accumulates_string => {
        "local s = \"\"\nfor i = 1, 4 do s = s .. i end\nprint(s)\n",
        "1234"
    },
    pairs_sum_hash_values => {
        "local t = {a = 2, b = 3}\nlocal s = 0\nfor _, v in pairs(t) do s = s + v end\nprint(s)\n",
        "5"
    },
    ipairs_stops_on_nil_slot => {
        "local t = {1, nil, 3}\nlocal c = 0\nfor _ in ipairs(t) do c = c + 1 end\nprint(c)\n",
        "1"
    },
    numeric_for_with_negative_step_to_zero => {
        "local last = nil\nfor i = 3, 1, -1 do last = i end\nprint(last)\n",
        "1"
    },
    break_in_pairs_loop_early => {
        "local t = {a=1, b=2, c=3}\nlocal n = 0\nfor _ in pairs(t) do\n  n = n + 1\n  if n == 2 then break end\nend\nprint(n)\n",
        "2"
    },
    repeat_until_reads_updated_local => {
        "local n = 0\nrepeat n = n + 2 until n >= 4\nprint(n)\n",
        "4"
    },
    ipairs_iterator_call_returns_index_and_value => {
        "local t = {10, 20}\nlocal i, v = ipairs(t)()\nprint(i .. \",\" .. v)\n",
        "1,10"
    },
    numeric_for_outer_local_unaffected_by_loop_var => {
        "local i = 99\nfor i = 1, 1 do end\nprint(i)\n",
        "99"
    },
    generic_for_triple_return_from_iterator => {
        "local function iter(s, k)\n  k = (k or 0) + 1\n  if k > #s then return nil end\n  return k, s[k], #s\nend\nlocal a, b, c = iter({\"x\"}, 0)\nprint(a .. b .. c)\n",
        "1x1"
    },
    while_loop_with_truthy_zero_runs_until_break => {
        "local n = 0\nwhile 0 do n = n + 1 if n == 2 then break end end\nprint(n)\n",
        "2"
    },
    pairs_visits_single_string_key_record => {
        "local t = {mode = \"rw\"}\nfor k in pairs(t) do print(k) end\n",
        "mode"
    },
    numeric_for_fractional_bounds_still_iterates => {
        "local last = 0\nfor i = 1.0, 3.0 do last = i end\nprint(last)\n",
        "3"
    },
    pairs_over_empty_table_runs_zero_iterations => {
        "local count = 0\nfor _ in pairs({}) do count = count + 1 end\nprint(count)\n",
        "0"
    },
    ipairs_ignores_hash_part_of_mixed_table => {
        "local t = {10, 20, x = 99}\nlocal sum = 0\nfor _, v in ipairs(t) do sum = sum + v end\nprint(sum)\n",
        "30"
    },
    numeric_for_single_iteration_range => {
        "local count = 0\nfor i = 5, 5 do count = count + 1 end\nprint(count)\n",
        "1"
    },
    next_called_manually_returns_pairs_in_sequence => {
        "local t = {a = 1}\nlocal k, v = next(t)\nprint(k .. tostring(v))\n",
        "a1"
    },
    ipairs_with_early_break_counts_correctly => {
        "local count = 0\nfor i, v in ipairs({10, 20, 30, 40, 50}) do\n  count = count + 1\n  if i == 3 then break end\nend\nprint(count)\n",
        "3"
    },
    next_on_empty_table_returns_nil => {
        "local k = next({})\nprint(tostring(k))\n",
        "nil"
    },
    numeric_for_large_range_with_accumulation => {
        "local sum = 0\nfor i = 1, 100 do sum = sum + i end\nprint(sum)\n",
        "5050"
    },
    iterator_state_encapsulated_in_closure => {
        "local function range(from, to)\n  local i = from - 1\n  return function()\n    i = i + 1\n    if i <= to then return i end\n  end\nend\nlocal t = {}\nfor v in range(3, 6) do t[#t+1] = v end\nprint(table.concat(t, ','))\n",
        "3,4,5,6"
    },
}
