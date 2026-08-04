-- vybe-test: lua/debug_locals/test_debug_getlocal_valid
-- origin: languages/lua/tests/lua/test_debug_locals.rs

local __w1 = "string 42"
local __i = 0

local function f() local a=42; local n, v = debug.getlocal(1, 1); do local __t = tostring(type(n)..' '..v); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end end; f()

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
