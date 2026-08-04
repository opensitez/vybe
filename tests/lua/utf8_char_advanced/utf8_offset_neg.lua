-- vybe-test: lua/utf8_char_advanced/utf8_offset_neg
-- origin: languages/lua/tests/lua/test_utf8_char_advanced.rs

local __w1 = "4"
local __i = 0

local s = "Aα😀"
do local __t = tostring(utf8.offset(s, -1, #s+1)); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
