-- vybe-test: lua/table_iteration_order/iter_pairs_all_hash
-- origin: languages/lua/tests/lua/test_table_iteration_order.rs

local __w1 = "2"
local __i = 0

local t = {x=1, y=2}
local count = 0
for k, v in pairs(t) do count = count + 1 end
do local __t = tostring(count); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
