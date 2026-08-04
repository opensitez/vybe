-- vybe-test: lua/assignment/global_assignment_from_function
-- origin: languages/lua/tests/lua/test_assignment.rs

local __w1 = "11"
local __i = 0

function setg() gmark = 11 end
setg()
do local __t = tostring(gmark); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
