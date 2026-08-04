-- vybe-test: lua/tables_mixed/test_mixed_table_constructor_overwrite_reverse
-- origin: languages/lua/tests/lua/test_tables_mixed.rs

local __w1 = "10"
local __i = 0

local t={[1]=20, 10}; do local __t = tostring(t[1]); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
