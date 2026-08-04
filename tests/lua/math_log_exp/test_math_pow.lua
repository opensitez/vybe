-- vybe-test: lua/math_log_exp/test_math_pow
-- origin: languages/lua/tests/lua/test_math_log_exp.rs

local __w1 = "8"
local __i = 0

do local __t = tostring(math.floor(math.pow(2, 3))); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
