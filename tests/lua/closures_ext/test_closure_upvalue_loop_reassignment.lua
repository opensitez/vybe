-- vybe-test: lua/closures_ext/test_closure_upvalue_loop_reassignment
-- origin: languages/lua/tests/lua/test_closures_ext.rs

local __w1 = "444"
local __i = 0

local t={}; local a=1; for i=1,3 do t[i]=function() return a end; a=a+1 end; do local __t = tostring(t[1]()..t[2]()..t[3]()); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
