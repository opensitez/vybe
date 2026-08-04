-- vybe-test: lua/table_iteration_order/iter_ipairs_stops_nil
-- origin: languages/lua/tests/lua/test_table_iteration_order.rs

local __w1 = "10,20"
local __i = 0

local t = {10, 20, nil, 40}
local values = {}
for i, v in ipairs(t) do values[#values+1] = v end
do local __t = tostring(table.concat(values, ",")); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
