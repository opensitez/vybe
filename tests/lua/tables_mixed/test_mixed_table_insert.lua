-- vybe-test: lua/tables_mixed/test_mixed_table_insert
-- origin: languages/lua/tests/lua/test_tables_mixed.rs

local __w1 = "10 1"
local __i = 0

local t={a=1}; table.insert(t, 10); do local __t = tostring(t[1]..' '..t.a); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
