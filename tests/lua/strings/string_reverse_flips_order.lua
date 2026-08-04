-- vybe-test: lua/strings/string_reverse_flips_order
-- origin: languages/lua/tests/lua/test_strings.rs

local __w1 = "aul"
local __i = 0

do local __t = tostring(string.reverse("lua")); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
