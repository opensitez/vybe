-- vybe-test: lua/tables_move/test_move_overlap_forward
-- origin: languages/lua/tests/lua/test_tables_move.rs

local __w1 = "1123"
local __i = 0

local t={1,2,3,4}; table.move(t, 1, 3, 2); do local __t = tostring(t[1]..t[2]..t[3]..t[4]); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
