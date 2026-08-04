-- vybe-test: lua/math_library/math_randomseed_then_random_deterministic_range
-- origin: languages/lua/tests/lua/test_math_library.rs

local __w1 = "true"
local __i = 0

math.randomseed(99)
local a = math.random(1, 3)
math.randomseed(99)
local b = math.random(1, 3)
do local __t = tostring(a == b); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
