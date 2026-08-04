-- vybe-test: lua/closures_nested/test_nested_closure_deep_mixed
-- origin: languages/lua/tests/lua/test_closures_nested.rs

local __w1 = "3132"
local __i = 0

local function f1() local a=10; return function() local b=20; return function() a=a+1; return a+b end end end; local f3 = f1()(); do local __t = tostring(f3()..f3()); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
