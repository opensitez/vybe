-- vybe-test: lua/debug_locals/debug_getlocal_on_c_function_returns_nil
-- origin: languages/lua/tests/lua/test_debug_locals.rs

local __w1 = "x 42"
local __i = 0

local name, val = debug.getlocal(print, 1)
do local __t = tostring(tostring(name) .. " " .. tostring(val)); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
