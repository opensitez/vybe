-- vybe-test: lua/assignment/assignment_evaluates_rhs_before_lhs
-- origin: languages/lua/tests/lua/test_assignment.rs

local __w1 = "2,2"
local __i = 0

local t = {1, 2, 3}
local i = 1
t[i], i = t[i + 1], i + 1
do local __t = tostring(t[1] .. "," .. i); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
