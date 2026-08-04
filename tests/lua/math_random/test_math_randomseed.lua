-- vybe-test: lua/math_random/test_math_randomseed
-- origin: languages/lua/tests/lua/test_math_random.rs

local __w1 = "true"
local __i = 0

math.randomseed(42); local r1 = math.random(); math.randomseed(42); local r2 = math.random(); do local __t = tostring(tostring(r1 == r2)); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
