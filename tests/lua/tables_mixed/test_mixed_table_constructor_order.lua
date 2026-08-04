-- vybe-test: lua/tables_mixed/test_mixed_table_constructor_order
-- origin: languages/lua/tests/lua/test_tables_mixed.rs

local __w1 = "10 20 1 2"
local __i = 0

local t={a=1, 10, b=2, 20}; do local __t = tostring(t[1]..' '..t[2]..' '..t.a..' '..t.b); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
