-- vybe-test: lua/loops_for_float_bounds/test_loops_for_float_bounds_decimal_stride_skips_fractionals
-- origin: languages/lua/tests/lua/test_loops_for_float_bounds.rs

local __w1 = "4"
local __i = 0

local count = 0
for i = 0.0, 2.0, 0.5 do if i > 0 then count = count + 1 end end
do local __t = tostring(count); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
