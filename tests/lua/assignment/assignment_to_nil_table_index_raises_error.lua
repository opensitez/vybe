-- vybe-test: lua/assignment/assignment_to_nil_table_index_raises_error
-- origin: languages/lua/tests/lua/test_assignment.rs

local __w1 = "false"
local __i = 0

local ok, err = pcall(function() local t; t.x = 1 end)
do local __t = tostring(ok); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
