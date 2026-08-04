-- vybe-test: lua/utf8_char_advanced/utf8_codepoint_basic
-- origin: languages/lua/tests/lua/test_utf8_char_advanced.rs

local __w1 = "945"
local __i = 0

local s = "α"
do local __t = tostring(utf8.codepoint(s)); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
