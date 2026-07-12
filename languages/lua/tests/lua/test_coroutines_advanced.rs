lua_print! {
    test_coroutine_ping_pong => {
        "local ping, pong
ping = coroutine.create(function(n)
    local sum = 0
    for i = 1, n do
        sum = sum + 1
        coroutine.resume(pong)
    end
    return sum
end)
pong = coroutine.create(function()
    while true do
        coroutine.yield()
    end
end)
local ok, res = coroutine.resume(ping, 3)
print(res)",
        "3"
    },
    test_coroutine_iterator_pattern => {
        "local function permgen(a, n)
    n = n or #a
    if n <= 1 then
        coroutine.yield(a)
    else
        for i = 1, n do
            a[n], a[i] = a[i], a[n]
            permgen(a, n - 1)
            a[n], a[i] = a[i], a[n]
        end
    end
end
local function permutations(a)
    local co = coroutine.create(function() permgen(a) end)
    return function()
        local code, res = coroutine.resume(co)
        return res
    end
end
local count = 0
for p in permutations({1, 2, 3}) do
    count = count + 1
end
print(count)",
        "6"
    },
    test_coroutine_yielding_across_pcall => {
        "local co = coroutine.create(function()
    local ok, err = pcall(function()
        coroutine.yield(42)
        error('boom')
    end)
    return ok, err
end)
local ok, res = coroutine.resume(co)
local ok2, ok_inner, err = coroutine.resume(co)
print(res .. ' ' .. tostring(ok_inner) .. ' ' .. tostring(string.find(err, 'boom') ~= nil))",
        "42 false true"
    },
    test_coroutine_wrap_error_propagation => {
        "local f = coroutine.wrap(function()
    error('wrap error')
end)
local ok, err = pcall(f)
print(tostring(ok) .. ' ' .. tostring(string.find(err, 'wrap error') ~= nil))",
        "false true"
    },
    test_coroutine_resume_dead => {
        "local co = coroutine.create(function() return 1 end)
coroutine.resume(co)
local ok, err = coroutine.resume(co)
print(tostring(ok) .. ' ' .. tostring(string.find(err, 'dead coroutine') ~= nil))",
        "false true"
    }
}
