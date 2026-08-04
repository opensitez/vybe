-- vybe-test: lua/table_library_extended/table_move_elements
-- origin: languages/lua/tests/lua/test_table_library_extended.rs

local __w1 = "10,30"
local __i = 0

local a = {10, 20, 30}
local b = {}
table.move(a, 1, 3, 1, b)
do local __t = tostring(b[1] .. "," .. b[3]); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
