-- vybe-test: lua/programs/stack_push_pop_using_table
-- origin: languages/lua/tests/lua/test_programs.rs

local __w1 = "2"
local __i = 0

local st = {}
table.insert(st, 1)
table.insert(st, 2)
do local __t = tostring(table.remove(st)); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
