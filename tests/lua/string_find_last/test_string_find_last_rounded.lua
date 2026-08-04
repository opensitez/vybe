-- vybe-test: lua/string_find_last/test_string_find_last_rounded
-- origin: languages/lua/tests/lua/test_string_find_last.rs

local __w1 = "true"
local __i = 0

do local __t = tostring(string.find("x8 x8 x8 x8 x8 x8 x8 x8 x8 x8 z8", "z8", 1, true) == 31); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
