-- vybe-test: lua/utf8_char_advanced/utf8_char_points
-- origin: languages/lua/tests/lua/test_utf8_char_advanced.rs

local __w1 = "Aα😀"
local __i = 0

do local __t = tostring(utf8.char(65, 0x3B1, 0x1F600)); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
