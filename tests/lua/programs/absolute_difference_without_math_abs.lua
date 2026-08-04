-- vybe-test: lua/programs/absolute_difference_without_math_abs
-- origin: languages/lua/tests/lua/test_programs.rs

local __w1 = "7"
local __i = 0

local a, b = 3, 10
local d = a - b
if d < 0 then d = -d end
do local __t = tostring(d); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
