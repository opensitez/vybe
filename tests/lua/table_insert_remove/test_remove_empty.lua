-- vybe-test: lua/table_insert_remove/test_remove_empty
-- origin: languages/lua/tests/lua/test_table_insert_remove.rs

local __w1 = "nil"
local __i = 0

local t={}; local v = table.remove(t); do local __t = tostring(tostring(v)); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
