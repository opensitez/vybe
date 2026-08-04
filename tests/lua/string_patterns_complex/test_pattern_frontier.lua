-- vybe-test: lua/string_patterns_complex/test_pattern_frontier
-- origin: languages/lua/tests/lua/test_string_patterns_complex.rs

local __w1 = "quick"
local __i = 0

local s = 'the quick brown fox'
local m = string.match(s, '%f[%w]quick%f[%W]')
do local __t = tostring(m); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
