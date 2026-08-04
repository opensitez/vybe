-- vybe-test: lua/tables_mixed/test_mixed_table_length_operator_hole_in_middle
-- origin: languages/lua/tests/lua/test_tables_mixed.rs

local __w1 = "true"
local __i = 0

local t={1, nil, 3, a=10}; local len=#t; do local __t = tostring(len==1 or len==3); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
