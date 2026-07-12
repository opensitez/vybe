lua_print! {
    test_for_gen_ipairs => {
        "local s=''; for k,v in ipairs({'a','b','c'}) do s=s..k..v end; print(s)",
        "1a2b3c"
    },
    test_for_gen_pairs => {
        "local s=''; local t={a=1,b=2}; for k,v in pairs(t) do s=s..k..v end; -- Can't strictly assert string order due to hash, so just count
         local count=0; for _ in pairs(t) do count=count+1 end; print(count)",
        "2"
    },
    test_for_gen_custom_iterator_stateful => {
        "local function iter(max) local i=0; return function() i=i+1; if i<=max then return i, i*2 end end end;
         local s=''; for k,v in iter(3) do s=s..k..v end; print(s)",
        "122436"
    },
    test_for_gen_custom_iterator_stateless => {
        "local function iter(state, var) var=var+1; if var<=state then return var, var*10 end end;
         local s=''; for k,v in iter, 3, 0 do s=s..k..v end; print(s)",
        "110220330"
    },
    test_for_gen_break => {
        "local s=''; for k,v in ipairs({10,20,30,40}) do s=s..v; if k==2 then break end end; print(s)",
        "1020"
    },
    test_for_gen_local_scope => {
        "local k=99; local v=88; for k,v in ipairs({1}) do end; print(k..' '..v)",
        "99 88"
    },
    test_for_gen_closure_capture => {
        "local t={}; for k,v in ipairs({'a','b'}) do t[k]=function() return v end end; print(t[1]()..t[2]())",
        "ab"
    },
    test_for_gen_multiple_returns_from_iterator => {
        "local function it() local i=0; return function() i=i+1; if i<3 then return i, 'A', 'B' end end end;
         local s=''; for a,b,c in it() do s=s..a..b..c end; print(s)",
        "1AB2AB"
    },
    test_for_gen_eval_iterator_once => {
        "local c=0; local function get_iter() c=c+1; return ipairs({10,20}) end;
         local s=''; for k,v in get_iter() do s=s..v end; print(s..' '..c)",
        "1020 1"
    },
    generic_for_using_next_as_iterator => {
        "local t = {a=1, b=2, c=3}\nlocal count = 0\nfor k, v in next, t do count = count + v end\nprint(count)\n",
        "6"
    },
    generic_for_ipairs_stops_at_nil_hole => {
        "local t = {10, 20, nil, 40}\nlocal sum = 0\nfor _, v in ipairs(t) do sum = sum + v end\nprint(sum)\n",
        "30"
    },
    generic_for_iterator_returns_nil_to_stop => {
        "local function countdown(from, cur)\n  if cur == nil then cur = from end\n  if cur <= 0 then return nil end\n  return cur - 1, cur\nend\nlocal s = ''\nfor _, v in countdown, 3 do s = s .. v end\nprint(s)\n",
        "321"
    },
    generic_for_pcall_inside_body => {
        "local results = {}\nfor i, v in ipairs({1, 2, 'bad', 4}) do\n  local ok, n = pcall(function() return v + 0 end)\n  results[i] = ok\nend\nprint(table.concat(results, ','))\n",
        "true,true,false,true"
    },
    generic_for_sorted_keys_iteration => {
        "local t = {c=3, a=1, b=2}\nlocal keys = {}\nfor k in pairs(t) do keys[#keys+1] = k end\ntable.sort(keys)\nlocal s = ''\nfor _, k in ipairs(keys) do s = s .. k .. t[k] end\nprint(s)\n",
        "a1b2c3"
    },
    generic_for_coroutine_producer_consumer => {
        "local function producer(t)\n  return coroutine.wrap(function()\n    for _, v in ipairs(t) do coroutine.yield(v) end\n  end)\nend\nlocal s = ''\nfor v in producer({10, 20, 30}) do s = s .. v .. ',' end\nprint(s)\n",
        "10,20,30,"
    },
    generic_for_closure_captures_loop_control_vars => {
        "local fns = {}\nfor i, v in ipairs({'x', 'y', 'z'}) do\n  fns[i] = function() return i .. v end\nend\nprint(fns[1]() .. ',' .. fns[2]() .. ',' .. fns[3]())\n",
        "1x,2y,3z"
    },
    generic_for_three_value_protocol_used_directly => {
        "local function stateless(state, i)\n  i = i + 1\n  if i <= state then return i end\nend\nlocal sum = 0\nfor i in stateless, 5, 0 do sum = sum + i end\nprint(sum)\n",
        "15"
    },
}
