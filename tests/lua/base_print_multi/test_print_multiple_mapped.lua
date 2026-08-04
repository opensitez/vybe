-- vybe-test: lua/base_print_multi/test_print_multiple_mapped
-- origin: languages/lua/tests/lua/test_base_print_multi.rs

local __w1 = "x\ty\tnil"
local __i = 0

do local __t = tostring("x") .. "\t" .. tostring("y") .. "\t" .. tostring(nil); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
