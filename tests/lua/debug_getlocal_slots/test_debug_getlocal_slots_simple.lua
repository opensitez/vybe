-- vybe-test: lua/debug_getlocal_slots/test_debug_getlocal_slots_simple
-- origin: languages/lua/tests/lua/test_debug_getlocal_slots.rs

local __w1 = "true"
local __i = 0

local function f()
  local a = 2
  local b = 3
  return debug.getlocal(1, 1) ~= nil and debug.getlocal(1, 2) ~= nil
end
do local __t = tostring(f()); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
