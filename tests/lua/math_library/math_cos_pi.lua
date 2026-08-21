-- vybe-test: lua/math_library/math_cos_pi
-- origin: languages/lua/tests/lua/test_math_library.rs

local __w1 = "-1.0"
local __i = 0

do local __t = tostring(math.cos(math.pi)); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
