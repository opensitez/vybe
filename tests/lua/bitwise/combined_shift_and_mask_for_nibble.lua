-- vybe-test: lua/bitwise/combined_shift_and_mask_for_nibble
-- origin: languages/lua/tests/lua/test_bitwise.rs

local __w1 = "10"
local __i = 0

do local __t = tostring((0xAB >> 4) & 0xF); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
