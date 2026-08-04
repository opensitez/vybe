-- vybe-test: lua/closures_ext/test_closure_multiple_upvalues
-- origin: languages/lua/tests/lua/test_closures_ext.rs

local __w1 = "123"
local __i = 0

local a,b,c=1,2,3; local function f() return a..b..c end; do local __t = tostring(f()); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
