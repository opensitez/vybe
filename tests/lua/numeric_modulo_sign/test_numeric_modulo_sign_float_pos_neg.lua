-- vybe-test: lua/numeric_modulo_sign/test_numeric_modulo_sign_float_pos_neg
-- origin: languages/lua/tests/lua/test_numeric_modulo_sign.rs

local __w1 = "-0.5"
local __i = 0

do local __t = tostring(8.5 % -3); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
