-- vybe-test: lua/numeric_bitwise_or_xor/test_numeric_bitwise_or_xor_or_xor_combo
-- origin: languages/lua/tests/lua/test_numeric_bitwise_or_xor.rs

local __w1 = "13"
local __i = 0

do local __t = tostring((13 | 7) ~ 2); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
