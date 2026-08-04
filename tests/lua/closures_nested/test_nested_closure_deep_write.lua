-- vybe-test: lua/closures_nested/test_nested_closure_deep_write
-- origin: languages/lua/tests/lua/test_closures_nested.rs

local __w1 = "12"
local __i = 0

local function f1() local a=0; return function() return function() return function() a=a+1; return a end end end end; local f4 = f1()()(); do local __t = tostring(f4()..f4()); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
