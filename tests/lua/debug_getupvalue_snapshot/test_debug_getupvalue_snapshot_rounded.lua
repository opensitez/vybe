-- vybe-test: lua/debug_getupvalue_snapshot/test_debug_getupvalue_snapshot_rounded
-- origin: languages/lua/tests/lua/test_debug_getupvalue_snapshot.rs

local __w1 = "true"
local __i = 0

local up = 8
local function f()
  return up
end
local name, value = debug.getupvalue(f, 1)
do local __t = tostring(name == "up" and value == 8); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
