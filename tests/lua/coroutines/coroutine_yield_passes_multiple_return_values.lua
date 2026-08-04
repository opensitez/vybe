-- vybe-test: lua/coroutines/coroutine_yield_passes_multiple_return_values
-- origin: languages/lua/tests/lua/test_coroutines.rs

local __w1 = "6"
local __i = 0

local co = coroutine.create(function() return 1, 2, 3 end)
local _, a, b, c = coroutine.resume(co)
do local __t = tostring(a + b + c); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
