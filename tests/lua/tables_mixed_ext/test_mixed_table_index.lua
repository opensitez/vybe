-- vybe-test: lua/tables_mixed_ext/test_mixed_table_index
-- origin: languages/lua/tests/lua/test_tables_mixed_ext.rs

local __w1 = "1 2"
local __i = 0

local t={}; t[true] = 1; t[false] = 2; do local __t = tostring(t[true]..' '..t[false]); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
