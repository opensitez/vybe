-- vybe-test: lua/string_patterns/pattern_match_returns_nil_on_failure
-- origin: languages/lua/tests/lua/test_string_patterns.rs

local __w1 = "nil"
local __i = 0

do local __t = tostring(tostring(string.match("abc", "z+"))); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
