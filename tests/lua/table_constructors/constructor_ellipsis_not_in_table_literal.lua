-- vybe-test: lua/table_constructors/constructor_ellipsis_not_in_table_literal
-- origin: languages/lua/tests/lua/test_table_constructors.rs

local __w1 = "3"
local __i = 0

local parts = {1,2,3}
local t = {table.unpack(parts)}
do local __t = tostring(t[3]); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
