-- vybe-test: lua/tables_move/test_move_empty_range
-- origin: languages/lua/tests/lua/test_tables_move.rs

local __w1 = "nil"
local __i = 0

local t1={1,2,3}; local t2={}; table.move(t1, 2, 1, 1, t2); do local __t = tostring(tostring(t2[1])); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
