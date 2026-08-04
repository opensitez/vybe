-- vybe-test: lua/programs/median_of_three_values
-- origin: languages/lua/tests/lua/test_programs.rs

local __w1 = "2"
local __i = 0

local a, b, c = 3, 1, 2
if a > b then a, b = b, a end
if b > c then b, c = c, b end
if a > b then a, b = b, a end
do local __t = tostring(b); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
