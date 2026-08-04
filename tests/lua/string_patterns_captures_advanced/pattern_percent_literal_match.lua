-- vybe-test: lua/string_patterns_captures_advanced/pattern_percent_literal_match
-- origin: languages/lua/tests/lua/test_string_patterns_captures_advanced.rs

local __w1 = "%b"
local __i = 0

local cap = string.match("a%b", "a(%%b)")
do local __t = tostring(cap); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
