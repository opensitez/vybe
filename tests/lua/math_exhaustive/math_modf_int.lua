-- vybe-test: lua/math_exhaustive/math_modf_int
-- origin: languages/lua/tests/lua/test_math_exhaustive.rs

local __w1 = "42,0"
local __i = 0

local i, f = math.modf(42)
do local __t = tostring(i .. "," .. f); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
