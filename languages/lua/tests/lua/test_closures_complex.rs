lua_print! {
    test_closures_mutual_recursion => {
        "local is_even, is_odd;
is_even = function(n)
    if n == 0 then return true else return is_odd(n - 1) end
end
is_odd = function(n)
    if n == 0 then return false else return is_even(n - 1) end
end
print(tostring(is_even(10)) .. ' ' .. tostring(is_odd(10)))",
        "true false"
    },
    test_closures_sibling_upvalues => {
        "local function make_counter()
    local count = 0
    return function() count = count + 1 return count end,
           function() count = count - 1 return count end
end
local inc, dec = make_counter()
print(inc() .. ' ' .. inc() .. ' ' .. dec())",
        "1 2 1"
    },
    test_closures_deep_upvalue_capture => {
        "local function outer(x)
    local function inner(y)
        local function deepest(z)
            return x + y + z
        end
        return deepest
    end
    return inner
end
local f = outer(10)
local g = f(20)
print(g(30))",
        "60"
    },
    test_closures_captured_loop_variables => {
        "local funcs = {}
for i = 1, 3 do
    local v = i
    funcs[i] = function() return v end
end
print(funcs[1]() .. ' ' .. funcs[2]() .. ' ' .. funcs[3]())",
        "1 2 3"
    },
    test_closures_escaping_upvalues => {
        "local f
do
    local x = 42
    f = function() return x end
end
print(f())",
        "42"
    },
    test_closures_shadowing_upvalues => {
        "local x = 10
local function f()
    local x = x + 5
    return function() return x end
end
print(f()())",
        "15"
    },
    test_closures_environment_interaction => {
        "local x = 1
local function get_x() return x end
local _ENV = {x = 10, get_x = get_x}
print(get_x())",
        "1"
    },
    test_closures_vararg_capture => {
        "local function capture_varargs(...)
    local args = {...}
    return function(i) return args[i] end
end
local f = capture_varargs('a', 'b', 'c')
print(f(2))",
        "b"
    }
}
