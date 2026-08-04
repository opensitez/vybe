-- vybe-test: lua/loops_numeric_edge/numeric_for_product
-- origin: languages/lua/tests/lua/test_loops_numeric_edge.rs

local __w1 = "120"
local __i = 0

local p = 1
for i = 1, 5 do p = p * i end
do local __t = tostring(p); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
