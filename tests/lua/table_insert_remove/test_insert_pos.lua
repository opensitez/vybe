-- vybe-test: lua/table_insert_remove/test_insert_pos
-- origin: languages/lua/tests/lua/test_table_insert_remove.rs

local __w1 = "123"
local __i = 0

local t={1,3}; table.insert(t, 2, 2); do local __t = tostring(t[1]..t[2]..t[3]); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
