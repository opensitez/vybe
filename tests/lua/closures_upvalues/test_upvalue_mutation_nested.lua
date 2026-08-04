-- vybe-test: lua/closures_upvalues/test_upvalue_mutation_nested
-- origin: languages/lua/tests/lua/test_closures_upvalues.rs

local __w1 = "2334"
local __i = 0

local function f() local a=1; return function() local b=2; return function() a=a+1; b=b+1; return a..b end end end; local c=f()(); do local __t = tostring(c()..c()); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
