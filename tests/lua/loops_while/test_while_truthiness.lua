-- vybe-test: lua/loops_while/test_while_truthiness
-- origin: languages/lua/tests/lua/test_loops_while.rs

local __w1 = "321"
local __i = 0

local i=3; local s=''; while i do s=s..i; i=i-1; if i==0 then i=nil end end; do local __t = tostring(s); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
