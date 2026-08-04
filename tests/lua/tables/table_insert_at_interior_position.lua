-- vybe-test: lua/tables/table_insert_at_interior_position
-- origin: languages/lua/tests/lua/test_tables.rs

local __w1 = "1,2,3"
local __i = 0

local t = {1, 3}
table.insert(t, 2, 2)
do local __t = tostring(table.concat(t, ",")); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
