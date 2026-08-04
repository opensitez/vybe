-- vybe-test: lua/string_patterns_captures_advanced/pattern_zero_or_more
-- origin: languages/lua/tests/lua/test_string_patterns_captures_advanced.rs

local __w1 = "b"
local __i = 0

local cap = string.match("abc", "a(b*)c")
do local __t = tostring(cap); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
