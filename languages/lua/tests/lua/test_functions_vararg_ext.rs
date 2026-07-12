lua_print! {
    test_vararg_multiple_returns => { "local function f(...) return ... end; local a, b = f(1, 2); print(a..' '..b)", "1 2" },
    test_vararg_select_index => { "local function f(...) return select(2, ...) end; local a = f(1, 2, 3); print(a)", "2" },
    test_vararg_select_count => { "local function f(...) return select('#', ...) end; local c = f(1, nil, 3); print(c)", "3" },
    test_vararg_select_negative => { "local function f(...) return select(-1, ...) end; local a = f(1, 2, 3); print(a)", "3" },
    test_vararg_in_table => { "local function f(...) local t={...}; return t[1]..t[2] end; print(f('a', 'b'))", "ab" },
    test_vararg_mixed => { "local function f(a, ...) return a, select('#', ...) end; local x, c = f(1, 2, 3); print(x..' '..c)", "1 2" }
}
