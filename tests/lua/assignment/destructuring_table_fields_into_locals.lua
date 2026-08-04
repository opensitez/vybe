-- vybe-test: lua/assignment/destructuring_table_fields_into_locals
-- origin: languages/lua/tests/lua/test_assignment.rs

local __w1 = "3"
local __i = 0

local t = {x = 1, y = 2}
local a, b = t.x, t.y
do local __t = tostring(a + b); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
