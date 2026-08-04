-- vybe-test: lua/loops_while/test_while_false_truthiness
-- origin: languages/lua/tests/lua/test_loops_while.rs

local __w1 = "0"
local __i = 0

local flag=false; local cnt=0; while flag do cnt=cnt+1 end; do local __t = tostring(cnt); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
