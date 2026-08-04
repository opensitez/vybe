-- vybe-test: lua/loops_for_float_bounds/test_loops_for_float_bounds_local_step_variable
-- origin: languages/lua/tests/lua/test_loops_for_float_bounds.rs

local __w1 = "9"
local __i = 0

local sum = 0
local step = 2.0
for i = 1.0, 6.0, step do sum = sum + i end
do local __t = tostring(sum); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
