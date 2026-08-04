-- vybe-test: lua/string_super/str_sup_reverse_2
-- origin: languages/lua/tests/lua/test_string_super.rs

local __w1 = "ba"
local __i = 0

do local __t = tostring(string.reverse("ab")); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
