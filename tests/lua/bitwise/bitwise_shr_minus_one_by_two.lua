-- vybe-test: lua/bitwise/bitwise_shr_minus_one_by_two
-- origin: languages/lua/tests/lua/test_bitwise.rs

local __w1 = "-1"
local __i = 0

do local __t = tostring(-1 >> 2); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
