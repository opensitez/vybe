-- vybe-test: lua/string_reverse_basic/test_string_reverse_basic_nested
-- origin: languages/lua/tests/lua/test_string_reverse_basic.rs

local __w1 = "true"
local __i = 0

do local __t = tostring(string.reverse(string.sub("abcdefghijklmnopqrst", 1, 11)) == string.sub("abcdefghijklmnopqrst", 1, 11):reverse()); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
