-- vybe-test: lua/string_sub_negative/test_string_sub_negative_offset
-- origin: languages/lua/tests/lua/test_string_sub_negative.rs

local __w1 = "true"
local __i = 0

do local __t = tostring(string.sub("abcdefghijklmnopqrstuvwxyz", -9, -1) == string.sub("abcdefghijklmnopqrstuvwxyz", -9, -1)); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
