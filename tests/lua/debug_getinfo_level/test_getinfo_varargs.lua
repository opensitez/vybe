-- vybe-test: lua/debug_getinfo_level/test_getinfo_varargs
-- origin: languages/lua/tests/lua/test_debug_getinfo_level.rs

local __w1 = "true"
local __i = 0

local function f(a,b,...)
  return debug.getinfo(1, "u").nparams
end
do local __t = tostring(f(1,2,3) > 0); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
