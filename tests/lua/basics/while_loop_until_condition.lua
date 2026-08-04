-- vybe-test: lua/basics/while_loop_until_condition
-- origin: languages/lua/tests/lua/test_basics.rs

local __w1 = "4"
local __i = 0

local n = 1
while n < 4 do n = n + 1 end
do local __t = tostring(n); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
