-- vybe-test: lua/tables_mixed/test_mixed_table_pack
-- origin: languages/lua/tests/lua/test_tables_mixed.rs

local __w1 = "2 1 3"
local __i = 0

local t=table.pack(1, 2); t.a=3; do local __t = tostring(t.n..' '..t[1]..' '..t.a); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
