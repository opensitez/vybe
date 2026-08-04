-- vybe-test: lua/debug_upvalues/test_debug_upvalueid_same
-- origin: languages/lua/tests/lua/test_debug_upvalues.rs

local __w1 = "true"
local __i = 0

local a=1; local function f1() return a end; local id = debug.upvalueid(f1, 1); do local __t = tostring(type(id) == 'userdata'); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
