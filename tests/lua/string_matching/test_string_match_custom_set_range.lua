-- vybe-test: lua/string_matching/test_string_match_custom_set_range
-- origin: languages/lua/tests/lua/test_string_matching.rs

local __w1 = "a"
local __i = 0

do local __t = tostring((string.match('0 1 A B a b', '[a-z]'))); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
