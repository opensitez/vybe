-- vybe-test: lua/loops_while/test_while_condition_false_initially
-- origin: languages/lua/tests/lua/test_loops_while.rs

local __w1 = "10"
local __i = 0

local i=10; while i<5 do i=i+1 end; do local __t = tostring(i); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
