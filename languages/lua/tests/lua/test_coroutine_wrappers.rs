lua_print! {
    test_wrap_basic => {
        "local f = coroutine.wrap(function()
    return 10
end)
print(f())",
        "10"
    },
    test_wrap_yield_resume => {
        "local f = coroutine.wrap(function()
    coroutine.yield(1)
    coroutine.yield(2)
    return 3
end)
print(f() .. ' ' .. f() .. ' ' .. f())",
        "1 2 3"
    },
    test_wrap_args => {
        "local f = coroutine.wrap(function(a, b)
    local c = coroutine.yield(a + b)
    return c * 2
end)
local res1 = f(10, 20)
local res2 = f(5)
print(res1 .. ' ' .. res2)",
        "30 10"
    },
    test_wrap_error_propagation => {
        "local f = coroutine.wrap(function()
    error('wrap error')
end)
local ok, err = pcall(f)
print(tostring(ok) .. ' ' .. tostring(string.find(err, 'wrap error') ~= nil))",
        "false true"
    },
    test_wrap_iterator => {
        "local function traverse(t)
    return coroutine.wrap(function()
        for i, v in ipairs(t) do
            coroutine.yield(v)
        end
    end)
end
local s = ''
for v in traverse({1, 2, 3}) do
    s = s .. v
end
print(s)",
        "123"
    }
}
