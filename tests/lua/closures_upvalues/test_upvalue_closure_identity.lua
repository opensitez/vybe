-- vybe-test: lua/closures_upvalues/test_upvalue_closure_identity
-- origin: languages/lua/tests/lua/test_closures_upvalues.rs

local __w1 = "false"
local __i = 0

local function f() local a=1; return function() return a end end; local c1=f(); local c2=f(); do local __t = tostring(tostring(c1==c2)); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
