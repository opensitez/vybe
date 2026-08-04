-- vybe-test: lua/coroutine_yield_receives/test_coroutine_yield_receives_prefixed
-- origin: languages/lua/tests/lua/test_coroutine_yield_receives.rs

local __w1 = "true"
local __i = 0

local c = coroutine.create(function(x)
  local y = coroutine.yield(x)
  return y == x * 2
end)
local _, first = coroutine.resume(c, 6)
local _, second = coroutine.resume(c, 12)
do local __t = tostring(first == 6 and second == true); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
