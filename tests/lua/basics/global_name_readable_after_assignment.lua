-- vybe-test: lua/basics/global_name_readable_after_assignment
-- origin: languages/lua/tests/lua/test_basics.rs

local __w1 = "10"
local __i = 0

score = 10
do local __t = tostring(score); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
