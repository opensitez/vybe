-- vybe-test: lua/assignment/multiple_assignment_order_lhs_resolved_before_rhs_written
-- origin: languages/lua/tests/lua/test_assignment.rs

local __w1 = "99"
local __i = 0

local t = {x = 1}
local a = t
t, a.x = {}, 99
do local __t = tostring(a.x); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
