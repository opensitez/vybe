-- vybe-test: lua/debug_locals/debug_getlocal_for_varargs
-- origin: languages/lua/tests/lua/test_debug_locals.rs

local __w1 = "(*vararg) 99"
local __i = 0

local function f(...)
  local name, val = debug.getlocal(1, -1)
  do local __t = tostring(tostring(name) .. " " .. tostring(val)); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end
end
f(99)

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
