-- vybe-test: lua/utf8_ext/test_utf8_char_max
-- origin: languages/lua/tests/lua/test_utf8_ext.rs

local __w1 = "1"
local __i = 0

local c = utf8.char(0x10FFFF); do local __t = tostring(utf8.len(c)); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
