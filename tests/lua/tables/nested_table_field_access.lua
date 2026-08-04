-- vybe-test: lua/tables/nested_table_field_access
-- origin: languages/lua/tests/lua/test_tables.rs

local __w1 = "8"
local __i = 0

local t = { inner = { value = 8 } }
do local __t = tostring(t.inner.value); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
