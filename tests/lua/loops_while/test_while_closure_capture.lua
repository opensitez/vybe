-- vybe-test: lua/loops_while/test_while_closure_capture
-- origin: languages/lua/tests/lua/test_loops_while.rs

local __w1 = "1"
local __i = 0

local f; local i=1; while i<=2 do local j=i; if i==1 then f = function() return j end end; i=i+1 end; do local __t = tostring(f()); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
