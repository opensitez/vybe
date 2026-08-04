-- vybe-test: lua/string_sub_len/test_sub_negative_start
-- origin: languages/lua/tests/lua/test_string_sub_len.rs

local __w1 = "llo"
local __i = 0

do local __t = tostring(string.sub('hello', -3)); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
