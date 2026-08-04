-- vybe-test: lua/debug_setlocal_change/test_debug_setlocal_change_guarded
-- origin: languages/lua/tests/lua/test_debug_setlocal_change.rs

local __w1 = "true"
local __i = 0

local function f()
  local x = 22
  debug.setlocal(1, 1, 23)
  do local __t = tostring(x == 23); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end
end
f()

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
