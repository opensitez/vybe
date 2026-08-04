-- vybe-test: lua/base_pairs_order/test_pairs_order_guarded
-- origin: languages/lua/tests/lua/test_base_pairs_order.rs

local __w1 = "true"
local __i = 0

local t = {}
for i = 1, 14 do table.insert(t, i) end
local total = 0
for _, v in ipairs(t) do total = total + v end
do local __t = tostring(total == (14 * (14 + 1)) / 2); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
