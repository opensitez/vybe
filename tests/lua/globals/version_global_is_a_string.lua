-- vybe-test: lua/globals/version_global_is_a_string
-- origin: languages/lua/tests/lua/test_globals.rs

local __w1 = "string"
local __i = 0

do local __t = tostring(type(_VERSION)); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
