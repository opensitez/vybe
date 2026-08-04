-- vybe-test: lua/coroutine_wrappers/test_wrap_args
-- origin: languages/lua/tests/lua/test_coroutine_wrappers.rs

local __w1 = "30 10"
local __i = 0

local f = coroutine.wrap(function(a, b)
    local c = coroutine.yield(a + b)
    return c * 2
end)
local res1 = f(10, 20)
local res2 = f(5)
do local __t = tostring(res1 .. ' ' .. res2); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
