-- vybe-test: lua/math_random/test_math_random_no_args
-- origin: languages/lua/tests/lua/test_math_random.rs

local __w1 = "true"
local __i = 0

local r = math.random(); do local __t = tostring(type(r) == 'number' and r >= 0 and r < 1); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
