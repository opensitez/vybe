-- vybe-test: lua/iteration/while_iterator_manual_index
-- origin: languages/lua/tests/lua/test_iteration.rs

local __w1 = "15"
local __i = 0

local t = {4, 5, 6}
local i, sum = 1, 0
while t[i] do sum = sum + t[i] i = i + 1 end
do local __t = tostring(sum); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
