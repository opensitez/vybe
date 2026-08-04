-- vybe-test: lua/string_patterns_exhaustive/pattern_class_x
-- origin: languages/lua/tests/lua/test_string_patterns_exhaustive.rs

local __w1 = "F"
local __i = 0

do local __t = tostring(string.match("g F g", "%x")); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
