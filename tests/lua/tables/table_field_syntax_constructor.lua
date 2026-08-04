-- vybe-test: lua/tables/table_field_syntax_constructor
-- origin: languages/lua/tests/lua/test_tables.rs

local __w1 = "3"
local __i = 0

local t = {x = 1, y = 2}
do local __t = tostring(t.x + t.y); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
