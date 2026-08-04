-- vybe-test: lua/tables/table_move_copies_slice_to_destination
-- origin: languages/lua/tests/lua/test_tables.rs

local __w1 = "2,3,3,4"
local __i = 0

local a = {1, 2, 3, 4}
table.move(a, 2, 3, 1, a)
do local __t = tostring(table.concat(a, ",")); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
