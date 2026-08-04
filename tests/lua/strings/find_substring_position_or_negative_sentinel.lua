-- vybe-test: lua/strings/find_substring_position_or_negative_sentinel
-- origin: languages/lua/tests/lua/test_strings.rs

local __w1 = "true"
local __i = 0

local i = string.find("hello", "z")
do local __t = tostring(i == nil); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
