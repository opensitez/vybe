-- vybe-test: lua/debug_upvalues/test_debug_getupvalue_valid
-- origin: languages/lua/tests/lua/test_debug_upvalues.rs

local __w1 = "string 42"
local __i = 0

local a=42; local function f() return a end; local name, val = debug.getupvalue(f, 1); do local __t = tostring(type(name)..' '..val); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
