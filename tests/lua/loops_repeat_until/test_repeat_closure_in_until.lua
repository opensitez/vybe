-- vybe-test: lua/loops_repeat_until/test_repeat_closure_in_until
-- origin: languages/lua/tests/lua/test_loops_repeat_until.rs

local __w1 = "4"
local __i = 0

local i=1; repeat local j=i; local f=function() return j>2 end; i=i+1 until f(); do local __t = tostring(i); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
