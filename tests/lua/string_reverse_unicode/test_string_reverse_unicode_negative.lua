-- vybe-test: lua/string_reverse_unicode/test_string_reverse_unicode_negative
-- origin: languages/lua/tests/lua/test_string_reverse_unicode.rs

local __w1 = "true"
local __i = 0

local s = "a6b"
do local __t = tostring(string.reverse(s) == "b6a"); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
