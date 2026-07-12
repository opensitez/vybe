lua_print! {
    test_env_load_with_env => { "local env = {a=10}; local f = load('return a', '', 't', env); print(f())", "10" },
    test_env_upvalue_is_env => { "local a=1; local function f() return a end; local name = debug.getupvalue(f, 1); print(name)", "_ENV" },
    test_env_lexical_override => { "local a=1; local _ENV={a=2}; print(a)", "1" },
    test_env_lexical_override_return => { "local a=1; local _ENV={a=2}; local function f() return a end; print(f())", "1" },
    test_env_lexical_override_no_print => { "local _ENV={a=2}; return a", "2" },
    test_env_lexical_shadow_eval => { "local a=1; local function f() local _ENV={}; return a end; print(f())", "1" },
    test_env_lexical_access_eval => { "local function f() local _ENV={b=2}; return b end; print(f())", "2" }
}
