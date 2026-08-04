-- vybe-test: lua/coroutines_advanced/test_coroutine_ping_pong
-- origin: languages/lua/tests/lua/test_coroutines_advanced.rs

local __w1 = "3"
local __i = 0

local ping, pong
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
do local __t = tostring(res); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
