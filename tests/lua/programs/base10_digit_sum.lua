-- vybe-test: lua/programs/base10_digit_sum
-- origin: languages/lua/tests/lua/test_programs.rs

local __w1 = "15"
local __i = 0

local function digit_sum(n)
  local s = 0
  while n > 0 do s = s + (n % 10); n = n // 10 end
  return s
end
do local __t = tostring(digit_sum(12345)); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
