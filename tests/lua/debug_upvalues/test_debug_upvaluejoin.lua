-- vybe-test: lua/debug_upvalues/test_debug_upvaluejoin
-- origin: languages/lua/tests/lua/test_debug_upvalues.rs

local __w1 = "1"
local __i = 0

local a=1; local b=2; local function f1() return a end; local function f2() return b end; debug.upvaluejoin(f1, 1, f2, 1); do local __t = tostring(f1()); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
