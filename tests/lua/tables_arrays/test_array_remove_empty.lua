-- vybe-test: lua/tables_arrays/test_array_remove_empty
-- origin: languages/lua/tests/lua/test_tables_arrays.rs

local __w1 = "nil 0"
local __i = 0

local t={}; local v=table.remove(t); do local __t = tostring(tostring(v)..' '..#t); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
