-- vybe-test: lua/numeric_hex_and_exponent/test_numeric_hex_and_exponent_mod_with_scientific
-- origin: languages/lua/tests/lua/test_numeric_hex_and_exponent.rs

local __w1 = "5.0"
local __i = 0

do local __t = tostring(25 % 1e1); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
