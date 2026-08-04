-- vybe-test: lua/closures_upvalues/test_upvalue_shared_between_closures
-- origin: languages/lua/tests/lua/test_closures_upvalues.rs

local __w1 = "245"
local __i = 0

local function outer() local a=1; local function f1() a=a+1; return a end; local function f2() a=a*2; return a end; return f1, f2 end; local f1, f2 = outer(); do local __t = tostring(f1()..f2()..f1()); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
