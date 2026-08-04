-- vybe-test: lua/table_constructors/constructor_list_then_record
-- origin: languages/lua/tests/lua/test_table_constructors.rs

local __w1 = "30"
local __i = 0

local t = {10, x=20}
do local __t = tostring(t[1]+t.x); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
