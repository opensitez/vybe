-- vybe-test: lua/tables_mixed/test_mixed_table_length_operator_no_holes
-- origin: languages/lua/tests/lua/test_tables_mixed.rs

local __w1 = "3"
local __i = 0

local t={1, 2, 3, a=10, b=20}; do local __t = tostring(#t); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
