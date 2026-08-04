-- vybe-test: lua/ipairs_pairs_edge/ipairs_sum
-- origin: languages/lua/tests/lua/test_ipairs_pairs_edge.rs

local __w1 = "30"
local __i = 0

local sum = 0
for _, v in ipairs({5, 10, 15}) do sum = sum + v end
do local __t = tostring(sum); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
