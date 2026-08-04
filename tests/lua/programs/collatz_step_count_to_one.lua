-- vybe-test: lua/programs/collatz_step_count_to_one
-- origin: languages/lua/tests/lua/test_programs.rs

local __w1 = "6"
local __i = 0

local n = 10
local steps = 0
while n ~= 1 do
  if n % 2 == 0 then n = n // 2 else n = 3 * n + 1 end
  steps = steps + 1
end
do local __t = tostring(steps); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
