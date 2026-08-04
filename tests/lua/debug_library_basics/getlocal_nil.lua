-- vybe-test: lua/debug_library_basics/getlocal_nil
-- origin: languages/lua/tests/lua/test_debug_library_basics.rs

local __w1 = "nil"
local __i = 0

local name, val = debug.getlocal(1, 100)
do local __t = tostring(tostring(name)); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
