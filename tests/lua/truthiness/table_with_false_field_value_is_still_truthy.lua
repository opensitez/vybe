-- vybe-test: lua/truthiness/table_with_false_field_value_is_still_truthy
-- origin: languages/lua/tests/lua/test_truthiness.rs

local __w1 = "table"
local __i = 0

local t = {flag = false}
if t then do local __t = tostring("table"); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
