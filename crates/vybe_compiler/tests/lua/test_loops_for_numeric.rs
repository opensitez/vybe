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
        "1.0,1.5,2.0,"
    },
    test_for_num_break => {
        "local s=''; for i=1,10 do s=s..i; if i==3 then break end end; print(s)",
        "123"
    },
    test_for_num_closure_capture => {
        "local t={}; for i=1,3 do t[i]=function() return i end end; print(t[1]()..t[2]()..t[3]())",
        "123"
    },
    test_for_num_expressions_as_bounds => {
        "local s=''; local a=1; local b=3; for i=a+1, b*2, b-1 do s=s..i end; print(s)",
        "246"
    }
}
