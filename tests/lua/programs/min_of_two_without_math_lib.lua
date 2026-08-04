-- vybe-test: lua/programs/min_of_two_without_math_lib
-- origin: languages/lua/tests/lua/test_programs.rs

local __w1 = "3"
local __i = 0

local a, b = 3, 7
do local __t = tostring(a < b and a or b); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
