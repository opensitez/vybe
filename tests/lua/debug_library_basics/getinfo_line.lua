-- vybe-test: lua/debug_library_basics/getinfo_line
-- origin: languages/lua/tests/lua/test_debug_library_basics.rs

local __w1 = "true"
local __i = 0

local info = debug.getinfo(1, "l")
do local __t = tostring(info.currentline > 0); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
