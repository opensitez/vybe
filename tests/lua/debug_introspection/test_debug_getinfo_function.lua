-- vybe-test: lua/debug_introspection/test_debug_getinfo_function
-- origin: languages/lua/tests/lua/test_debug_introspection.rs

local __w1 = "table"
local __i = 0

local info = debug.getinfo(print); do local __t = tostring(type(info)); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
