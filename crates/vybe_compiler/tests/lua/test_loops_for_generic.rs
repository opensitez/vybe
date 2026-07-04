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
    }
}
