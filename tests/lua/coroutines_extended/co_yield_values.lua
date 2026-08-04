-- vybe-test: lua/coroutines_extended/co_yield_values
-- origin: languages/lua/tests/lua/test_coroutines_extended.rs

local __w1 = "a\tb"
local __i = 0

local co = coroutine.create(function() coroutine.yield("a", "b") end)
local _, r1, r2 = coroutine.resume(co)
do local __t = tostring(r1) .. "\t" .. tostring(r2); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
