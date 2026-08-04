-- vybe-test: lua/tables/use_table_as_stack_with_push_pop
-- origin: languages/lua/tests/lua/test_tables.rs

local __w1 = "20"
local __i = 0

local st = {}
table.insert(st, 10)
table.insert(st, 20)
do local __t = tostring(table.remove(st)); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
