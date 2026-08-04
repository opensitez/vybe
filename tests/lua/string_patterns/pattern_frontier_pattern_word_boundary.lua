-- vybe-test: lua/string_patterns/pattern_frontier_pattern_word_boundary
-- origin: languages/lua/tests/lua/test_string_patterns.rs

local __w1 = "w"
local __i = 0

do local __t = tostring(string.match("word", "%f[%a]w")); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
