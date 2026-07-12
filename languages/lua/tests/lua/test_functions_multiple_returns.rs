lua_print! {
    test_multret_basic => { "local function f() return 1, 2, 3 end; local a,b,c = f(); print(a..b..c)", "123" },
    test_multret_truncation_assignment => { "local function f() return 1, 2, 3 end; local a = f(); print(a)", "1" },
    test_multret_truncation_middle_assignment => { "local function f() return 1, 2, 3 end; local a, b = f(), 4; print(a..b)", "14" },
    test_multret_padding_assignment => { "local function f() return 1 end; local a, b = f(); print(a..' '..(b or 'nil'))", "1 nil" },
    test_multret_trailing_in_function_call => { "local function f() return 2, 3 end; local function g(a,b,c) return a..b..c end; print(g(1, f()))", "123" },
    test_multret_middle_in_function_call => { "local function f() return 1, 2 end; local function g(a,b,c) return a..b..tostring(c) end; print(g(f(), 3))", "13nil" },
    test_multret_trailing_in_table_constructor => { "local function f() return 2, 3 end; local t = {1, f()}; print(t[1]..t[2]..t[3])", "123" },
    test_multret_middle_in_table_constructor => { "local function f() return 1, 2 end; local t = {f(), 3}; print(t[1]..t[2]..(t[3] or 'nil'))", "13nil" },
    test_multret_trailing_in_return => { "local function f() return 2, 3 end; local function g() return 1, f() end; local a,b,c = g(); print(a..b..c)", "123" },
    test_multret_middle_in_return => { "local function f() return 1, 2 end; local function g() return f(), 3 end; local a,b,c = g(); print(a..b..tostring(c))", "13nil" },
    test_multret_parentheses_truncate => { "local function f() return 1, 2, 3 end; local a,b = (f()); print(a..' '..(b or 'nil'))", "1 nil" },
    test_multret_parentheses_in_call => { "local function f() return 1, 2 end; local function g(...) return select('#', ...) end; print(g((f())))", "1" },
    test_multret_unpack_trailing => { "local function g(...) return select('#', ...) end; print(g(1, table.unpack({2,3})))", "3" },
    test_multret_unpack_middle => { "local function g(...) return select('#', ...) end; print(g(table.unpack({1,2}), 3))", "2" },
    test_multret_logical_and_truncates => { "local function f() return 1, 2 end; local a, b = true and f(); print(a..' '..(b or 'nil'))", "1 2" },
    test_multret_logical_or_truncates => { "local function f() return false, 2 end; local a, b = f() or 3; print(a..' '..(b or 'nil'))", "3 nil" }
}
