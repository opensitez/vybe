-- vybe-test: lua/programs/running_sum_with_numeric_for
-- origin: languages/lua/tests/lua/test_programs.rs

local __w1 = "15"
local __i = 0

local sum, n = 0, 5
for i = 1, n do sum = sum + i end
do local __t = tostring(sum); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
