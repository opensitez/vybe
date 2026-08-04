-- vybe-test: lua/closures_upvalues/test_upvalue_multiple_assignments
-- origin: languages/lua/tests/lua/test_closures_upvalues.rs

local __w1 = "12"
local __i = 0

local function f() local a,b=1,2; return function() return a..b end end; do local __t = tostring(f()()); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
