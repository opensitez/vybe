-- vybe-test: lua/programs/digit_sum_of_integer
-- origin: languages/lua/tests/lua/test_programs.rs

local __w1 = "6"
local __i = 0

local n = 123
local s = 0
while n > 0 do
  s = s + (n % 10)
  n = n // 10
end
do local __t = tostring(s); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
