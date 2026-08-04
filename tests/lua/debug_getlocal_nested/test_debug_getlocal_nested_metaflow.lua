-- vybe-test: lua/debug_getlocal_nested/test_debug_getlocal_nested_metaflow
-- origin: languages/lua/tests/lua/test_debug_getlocal_nested.rs

local __w1 = "true"
local __i = 0

local function inner()
  local z = 12
  return debug.getlocal(2, 1) ~= nil or debug.getlocal(1, 1) ~= nil
end
local function outer()
  local y = 13
  return inner()
end
do local __t = tostring(outer()); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
