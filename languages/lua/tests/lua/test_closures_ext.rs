lua_print! {
    test_closure_upvalue_loop_reassignment => { "local t={}; local a=1; for i=1,3 do t[i]=function() return a end; a=a+1 end; print(t[1]()..t[2]()..t[3]())", "444" },
    test_closure_upvalue_shadow_loop => { "local t={}; for i=1,3 do local a=i; t[i]=function() return a end end; print(t[1]()..t[2]()..t[3]())", "123" },
    test_closure_deep_recursion_upvalue => { "local function f(n) if n==0 then return function() return 42 end else return f(n-1) end end; print(f(10)())", "42" },
    test_closure_multiple_upvalues => { "local a,b,c=1,2,3; local function f() return a..b..c end; print(f())", "123" },
    test_closure_upvalue_table => { "local t={a=1}; local function f() t.a=t.a+1 return t.a end; f(); print(t.a)", "2" }
}
