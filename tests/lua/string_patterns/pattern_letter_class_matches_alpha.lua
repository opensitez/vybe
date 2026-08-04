-- vybe-test: lua/string_patterns/pattern_letter_class_matches_alpha
-- origin: languages/lua/tests/lua/test_string_patterns.rs

local __w1 = "abc"
local __i = 0

do local __t = tostring(string.match("1abc2", "%a+")); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
