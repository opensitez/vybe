-- vybe-test: lua/debug_getinfo_level/test_getinfo_not_variadic
-- origin: languages/lua/tests/lua/test_debug_getinfo_level.rs

local __w1 = "true"
local __i = 0

local function f(a,b)
  return debug.getinfo(1, "u").isvararg and 1 or 0
end
do local __t = tostring(type(debug.getinfo(1, "u").isvararg) == "boolean"); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
