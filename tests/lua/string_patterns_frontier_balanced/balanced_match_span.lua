-- vybe-test: lua/string_patterns_frontier_balanced/balanced_match_span
-- origin: languages/lua/tests/lua/test_string_patterns_frontier_balanced.rs

local __w1 = "(x+y)"
local __i = 0

local m = string.match("(x+y)", "%b()")
do local __t = tostring(m); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
