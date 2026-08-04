-- vybe-test: lua/utf8/utf8_codepoint_multiple_characters
-- origin: languages/lua/tests/lua/test_utf8.rs

local __w1 = "97,98"
local __i = 0

local c1, c2 = utf8.codepoint("ab", 1, 2)
do local __t = tostring(c1 .. "," .. c2); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
