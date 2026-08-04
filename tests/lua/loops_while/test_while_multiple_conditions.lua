-- vybe-test: lua/loops_while/test_while_multiple_conditions
-- origin: languages/lua/tests/lua/test_loops_while.rs

local __w1 = "1524"
local __i = 0

local a=1; local b=5; local s=''; while a<3 and b>3 do s=s..a..b; a=a+1; b=b-1 end; do local __t = tostring(s); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
