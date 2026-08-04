-- vybe-test: lua/string_matching/test_string_match_character_class_space
-- origin: languages/lua/tests/lua/test_string_matching.rs

local __w1 = " "
local __i = 0

do local __t = tostring((string.match('a b', '%s'))); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
