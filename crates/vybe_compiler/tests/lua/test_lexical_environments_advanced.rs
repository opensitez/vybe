lua_print! {
    test_env_shadowing => {
        "local _ENV = {print = print}
local a = 1
_ENV.a = 2
print(a)",
        "1"
    },
    test_env_dynamic_update => {
        "local env = {x = 10}
local f = load('return x', '', 't', env)
env.x = 20
print(f())",
        "20"
    },
    test_env_sandboxed_load => {
        "local f = load('return math and type(math) or \"nil\"', '', 't', {})
print(f())",
        "nil"
    },
    test_env_multiple_loads => {
        "local env = {}
local f1 = load('x = 5', '', 't', env)
local f2 = load('return x', '', 't', env)
f1()
print(f2())",
        "5"
    },
    test_env_change_upvalue => {
        "local _ENV = {print=print, type=type, debug=debug, load=load}
local function get_a() return a end
_ENV.a = 100
print(get_a())",
        "100"
    },
    test_env_set_upvalue_via_debug => {
        "local function f() return x end
local name, val = debug.getupvalue(f, 1)
if name == '_ENV' then
    debug.setupvalue(f, 1, {x = 42})
end
print(f())",
        "42"
    }
}
