-- vybe-test: lua/loops_for_float_bounds/test_loops_for_float_bounds_negative_bound_inclusive
-- origin: languages/lua/tests/lua/test_loops_for_float_bounds.rs

local __w1 = "3"
local __i = 0

local values = 0
for i = 0.0, -4.0, -2.0 do values = values + 1 end
do local __t = tostring(values); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
