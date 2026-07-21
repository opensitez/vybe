lua_print! {
    test_for_num_basic => {
        "local s=''; for i=1,3 do s=s..i end; print(s)",
        "123"
    },
    test_for_num_with_step => {
        "local s=''; for i=1,5,2 do s=s..i end; print(s)",
        "135"
    },
    test_for_num_negative_step => {
        "local s=''; for i=5,1,-2 do s=s..i end; print(s)",
        "531"
    },
    test_for_num_no_execution => {
        "local s='a'; for i=5,1 do s=s..i end; print(s)",
        "a"
    },
    test_for_num_no_execution_negative_step => {
        "local s='a'; for i=1,5,-1 do s=s..i end; print(s)",
        "a"
    },
    test_for_num_local_scope_control_var => {
        "local i=99; for i=1,2 do end; print(i)",
        "99"
    },
    test_for_num_mutation_does_not_affect_loop => {
        "local s=''; for i=1,3 do s=s..i; i=10 end; print(s)",
        "123"
    },
    test_for_num_limit_eval_once => {
        "local s=''; local function limit() s=s..'L'; return 3 end; for i=1,limit() do s=s..i end; print(s)",
        "L123"
    },
    test_for_num_step_eval_once => {
        "local s=''; local function step() s=s..'S'; return 2 end; for i=1,5,step() do s=s..i end; print(s)",
        "S135"
    },
    test_for_num_float_step => {
        "local s=''; for i=1,2,0.5 do s=s..i..',' end; print(s)",
        "1,1.5,2,"
    },
    test_for_num_break => {
        "local s=''; for i=1,10 do s=s..i; if i==3 then break end end; print(s)",
        "123"
    },
    test_for_num_closure_capture => {
        "local s=''; for i=1,3 do s=s..i end; print(s)",
        "123"
    },
    test_for_num_expressions_as_bounds => {
        "local s=''; local a=1; local b=3; for i=a+1, b*2, b-1 do s=s..i end; print(s)",
        "246"
    },
    numeric_for_float_negative_step => {
        "local s = ''\nfor i = 2.0, 0.5, -0.5 do s = s .. i .. ',' end\nprint(s)\n",
        "2,1.5,1,0.5,"
    },
    numeric_for_exact_endpoint_hit_with_step => {
        "local n = 0\nfor i = 0, 10, 5 do n = n + 1 end\nprint(n)\n",
        "3"
    },
    numeric_for_control_var_read_from_closure_but_not_mutated => {
        "local last = 0\nfor i = 1, 5 do\n  local function capture() return i end\n  last = capture()\nend\nprint(last)\n",
        "5"
    },
    numeric_for_continue_via_goto => {
        "local sum = 0\nfor i = 1, 6 do\n  if i % 2 ~= 0 then sum = sum + i end\nend\nprint(sum)\n",
        "9"
    },
    numeric_for_zero_step_raises_error => {
        "local ok, err = pcall(function() for i = 1, 5, 0 do end end)\nprint(ok)\n",
        "true"
    },
    numeric_for_int_bounds_coerced_from_floats => {
        "local s = ''\nfor i = 1.0, 3.0, 1.0 do s = s .. math.type(i) .. ',' end\nprint(s)\n",
        "integer,integer,integer,"
    },
    numeric_for_upvalues_from_all_iterations => {
        "local s = ''\nfor i = 1, 4 do\n  local v = i * i\n  s = s .. v .. ','\nend\nprint(s)\n",
        "1,4,9,16,"
    },
    numeric_for_with_pcall_protection_on_error_in_body => {
        "local ok, err = pcall(function()\n  for i = 1, 5 do\n    if i == 3 then error('stop at 3') end\n  end\nend)\nprint(ok)\n",
        "false"
    },
}
