-- vybe-test: lua/debug_getinfo_level/test_getinfo_defined_name
-- origin: languages/lua/tests/lua/test_debug_getinfo_level.rs

local __w1 = "inner"
local __i = 0

local function outer()
  local function inner() end
  local info = debug.getinfo(inner, "n")
  return info.name
end
do local __t = tostring(outer()); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
