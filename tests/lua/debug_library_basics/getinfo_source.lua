-- vybe-test: lua/debug_library_basics/getinfo_source
-- origin: languages/lua/tests/lua/test_debug_library_basics.rs

local __w1 = "string"
local __i = 0

local info = debug.getinfo(1, "S")
do local __t = tostring(type(info.source)); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
