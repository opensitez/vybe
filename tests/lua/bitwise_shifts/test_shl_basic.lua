-- vybe-test: lua/bitwise_shifts/test_shl_basic
-- origin: languages/lua/tests/lua/test_bitwise_shifts.rs

local __w1 = "4"
local __i = 0

do local __t = tostring(1 << 2); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
