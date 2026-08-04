-- vybe-test: lua/assignment/indexed_assignment_to_table_field
-- origin: languages/lua/tests/lua/test_assignment.rs

local __w1 = "7"
local __i = 0

local t = {}
t["k"] = 7
do local __t = tostring(t.k); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
