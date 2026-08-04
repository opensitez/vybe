-- vybe-test: lua/basics/assign_global_then_read_in_print
-- origin: languages/lua/tests/lua/test_basics.rs

local __w1 = "54"
local __i = 0

version = 54
do local __t = tostring(version); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
