-- vybe-test: lua/debug_getinfo_level/test_getinfo_negative_level
-- origin: languages/lua/tests/lua/test_debug_getinfo_level.rs

local __w1 = "true"
local __i = 0

local info = debug.getinfo(-1, "n")
do local __t = tostring(type(info) == "table" or type(info) == "nil"); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
