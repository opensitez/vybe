-- vybe-test: lua/loops_for_float_bounds/test_loops_for_float_bounds_zero_based_float_bounds
-- origin: languages/lua/tests/lua/test_loops_for_float_bounds.rs

local __w1 = "6"
local __i = 0

local sum = 0
for i = 0.0, 4.0, 2.0 do sum = sum + i end
do local __t = tostring(sum); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
