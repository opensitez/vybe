-- vybe-test: lua/programs/sum_of_squares
-- origin: languages/lua/tests/lua/test_programs.rs

local __w1 = "14"
local __i = 0

local t = {1, 2, 3}
local s = 0
for i = 1, #t do s = s + t[i] * t[i] end
do local __t = tostring(s); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
