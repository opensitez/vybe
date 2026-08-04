-- vybe-test: lua/coroutines_nested_yield/coroutine_status_during_yield
-- origin: languages/lua/tests/lua/test_coroutines_nested_yield.rs

local __w1 = "running"
local __i = 0

local co
co = coroutine.create(function()
  coroutine.yield(coroutine.status(co))
end)
local _, status = coroutine.resume(co)
do local __t = tostring(status); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
