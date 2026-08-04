-- vybe-test: lua/strings/replace_first_space_with_dash
-- origin: languages/lua/tests/lua/test_strings.rs

local __w1 = "a-b c\t1"
local __i = 0

do local __t = tostring(string.gsub("a b c", " ", "-", 1)); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
