-- vybe-test: lua/closures_upvalues/test_upvalue_basic
-- origin: languages/lua/tests/lua/test_closures_upvalues.rs

local __w1 = "1"
local __i = 0

local function outer() local a=1; return function() return a end end; do local __t = tostring(outer()()); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
