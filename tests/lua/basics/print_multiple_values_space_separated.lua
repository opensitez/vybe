-- vybe-test: lua/basics/print_multiple_values_space_separated
-- origin: languages/lua/tests/lua/test_basics.rs

local __w1 = "1\t2\t3"
local __i = 0

do local __t = tostring(1) .. "\t" .. tostring(2) .. "\t" .. tostring(3); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
