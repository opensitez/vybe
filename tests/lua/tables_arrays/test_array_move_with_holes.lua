-- vybe-test: lua/tables_arrays/test_array_move_with_holes
-- origin: languages/lua/tests/lua/test_tables_arrays.rs

local __w1 = "1 nil 3"
local __i = 0

local t1={1,nil,3}; local t2={}; table.move(t1, 1, 3, 1, t2); do local __t = tostring(t2[1]..' '..tostring(t2[2])..' '..t2[3]); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
