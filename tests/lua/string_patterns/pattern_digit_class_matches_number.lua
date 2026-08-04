-- vybe-test: lua/string_patterns/pattern_digit_class_matches_number
-- origin: languages/lua/tests/lua/test_string_patterns.rs

local __w1 = "9"
local __i = 0

do local __t = tostring(string.match("x9y", "%d")); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
