-- vybe-test: lua/tables/table_move_to_different_table
-- origin: languages/lua/tests/lua/test_tables.rs

local __w1 = "100,10,20"
local __i = 0

local t1 = {10, 20}
local t2 = {100, 200}
table.move(t1, 1, 2, 2, t2)
do local __t = tostring(table.concat(t2, ",")); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
