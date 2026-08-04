-- vybe-test: lua/string_matching/test_string_match_repetition_zero_or_more
-- origin: languages/lua/tests/lua/test_string_matching.rs

local __w1 = ">  <"
local __i = 0

local a, b = string.match('a  b', 'a(%s*)b'); do local __t = tostring('>'..a..'<'); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
