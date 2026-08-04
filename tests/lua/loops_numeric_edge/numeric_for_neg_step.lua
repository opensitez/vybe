-- vybe-test: lua/loops_numeric_edge/numeric_for_neg_step
-- origin: languages/lua/tests/lua/test_loops_numeric_edge.rs

local __w1 = "15"
local __i = 0

local s = 0
for i = 5, 1, -1 do s = s + i end
do local __t = tostring(s); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
