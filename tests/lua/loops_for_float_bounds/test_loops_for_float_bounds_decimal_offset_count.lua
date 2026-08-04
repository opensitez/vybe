-- vybe-test: lua/loops_for_float_bounds/test_loops_for_float_bounds_decimal_offset_count
-- origin: languages/lua/tests/lua/test_loops_for_float_bounds.rs

local __w1 = "5"
local __i = 0

local count = 0
for i = 0.0, 1.0, 0.25 do count = count + 1 end
do local __t = tostring(count); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
