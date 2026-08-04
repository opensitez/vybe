-- vybe-test: lua/table_iteration_order/iter_delete_current
-- origin: languages/lua/tests/lua/test_table_iteration_order.rs

local __w1 = "3"
local __i = 0

local t = {a=1, b=2, c=3}
local keys = {}
for k, v in pairs(t) do
  keys[#keys+1] = k
  t[k] = nil
end
do local __t = tostring(#keys); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
