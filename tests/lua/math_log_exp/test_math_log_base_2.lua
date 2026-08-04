-- vybe-test: lua/math_log_exp/test_math_log_base_2
-- origin: languages/lua/tests/lua/test_math_log_exp.rs

local __w1 = "3"
local __i = 0

do local __t = tostring(math.floor(math.log(8, 2) + 0.5)); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
