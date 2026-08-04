-- vybe-test: lua/iteration/numeric_for_with_negative_step_to_zero
-- origin: languages/lua/tests/lua/test_iteration.rs

local __w1 = "1"
local __i = 0

local last = nil
for i = 3, 1, -1 do last = i end
do local __t = tostring(last); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
