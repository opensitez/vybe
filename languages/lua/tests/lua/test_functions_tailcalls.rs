lua_print! {
    test_tailcall_basic => { "local function f() return 42 end; local function g() return f() end; print(g())", "42" },
    test_tailcall_multiple_returns => { "local function f() return 1, 2 end; local function g() return f() end; local a, b = g(); print(a..' '..b)", "1 2" },
    test_tailcall_deep => { "local function f(n) if n==0 then return 42 else return f(n-1) end end; print(f(100))", "42" },
    test_tailcall_not_tailcall_due_to_parentheses => { "local function f() return 42 end; local function g() return (f()) end; print(g())", "42" },
    test_tailcall_not_tailcall_due_to_operation => { "local function f() return 42 end; local function g() return f() + 0 end; print(g())", "42" },
    test_tailcall_mutually_recursive => { "local f, g; f = function(n) if n==0 then return 'f' else return g(n-1) end end; g = function(n) if n==0 then return 'g' else return f(n-1) end end; print(f(9)..' '..f(10))", "g f" }
}
