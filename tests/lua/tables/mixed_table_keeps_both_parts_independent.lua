-- vybe-test: lua/tables/mixed_table_keeps_both_parts_independent
-- origin: languages/lua/tests/lua/test_tables.rs

local __w1 = "30"
local __i = 0

local t = {10, key = 20}
do local __t = tostring(t[1] + t.key); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
