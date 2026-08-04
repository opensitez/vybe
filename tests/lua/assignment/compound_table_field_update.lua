-- vybe-test: lua/assignment/compound_table_field_update
-- origin: languages/lua/tests/lua/test_assignment.rs

local __w1 = "3"
local __i = 0

local t = {n = 1}
t.n = t.n + 2
do local __t = tostring(t.n); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
