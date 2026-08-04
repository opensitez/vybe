-- vybe-test: lua/debug_upvalues/test_debug_getupvalue_invalid_index
-- origin: languages/lua/tests/lua/test_debug_upvalues.rs

local __w1 = "nil nil"
local __i = 0

local function f() return 1 end; local name, val = debug.getupvalue(f, 2); do local __t = tostring(tostring(name)..' '..tostring(val)); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
