-- vybe-test: lua/programs/average_of_three_numbers
-- origin: languages/lua/tests/lua/test_programs.rs

local __w1 = "4"
local __i = 0

local a, b, c = 2, 4, 6
do local __t = tostring((a + b + c) / 3); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
