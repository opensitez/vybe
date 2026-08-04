-- vybe-test: lua/tables_mixed_ext/test_mixed_table_basic
-- origin: languages/lua/tests/lua/test_tables_mixed_ext.rs

local __w1 = "1 2 3 4"
local __i = 0

local t={1, 2, a=3, b=4}; do local __t = tostring(t[1]..' '..t[2]..' '..t.a..' '..t.b); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
