-- vybe-test: lua/iteration/numeric_for_large_range_with_accumulation
-- origin: languages/lua/tests/lua/test_iteration.rs

local __w1 = "5050"
local __i = 0

local sum = 0
for i = 1, 100 do sum = sum + i end
do local __t = tostring(sum); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
