-- vybe-test: lua/tables/table_move_overlapping_forward
-- origin: languages/lua/tests/lua/test_tables.rs

local __w1 = "20,30,30"
local __i = 0

local t = {10, 20, 30}
table.move(t, 2, 3, 1)
do local __t = tostring(table.concat(t, ",")); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
