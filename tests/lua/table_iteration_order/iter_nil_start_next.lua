-- vybe-test: lua/table_iteration_order/iter_nil_start_next
-- origin: languages/lua/tests/lua/test_table_iteration_order.rs

local __w1 = "a\t1"
local __i = 0

local t = {a=1}
local k, v = next(t, nil)
do local __t = tostring(k) .. "\t" .. tostring(v); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
