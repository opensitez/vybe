-- vybe-test: lua/debug_upvalues/debug_upvaluejoin_same_function_different_upvalues
-- origin: languages/lua/tests/lua/test_debug_upvalues.rs

local __w1 = "99,2"
local __i = 0

local a = 1; local b = 2
local function f() return a, b end
debug.upvaluejoin(f, 2, f, 1)
a = 99
do local __t = tostring(f()); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
