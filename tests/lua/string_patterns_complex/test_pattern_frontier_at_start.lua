-- vybe-test: lua/string_patterns_complex/test_pattern_frontier_at_start
-- origin: languages/lua/tests/lua/test_string_patterns_complex.rs

local __w1 = "the"
local __i = 0

local s = 'the quick'
local m = string.match(s, '%f[%w]the%f[%W]')
do local __t = tostring(m); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
