-- vybe-test: lua/string_exhaustive/str_exh_upper
-- origin: languages/lua/tests/lua/test_string_exhaustive.rs

local __w1 = "HELLO"
local __i = 0

do local __t = tostring(string.upper("hello")); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
