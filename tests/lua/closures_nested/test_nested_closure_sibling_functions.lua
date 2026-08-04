-- vybe-test: lua/closures_nested/test_nested_closure_sibling_functions
-- origin: languages/lua/tests/lua/test_closures_nested.rs

local __w1 = "1122"
local __i = 0

local function outer() local a=1; local function inner1() a=a+10; return a end; local function inner2() a=a*2; return a end; return function() return inner1()..inner2() end end; do local __t = tostring(outer()()); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
