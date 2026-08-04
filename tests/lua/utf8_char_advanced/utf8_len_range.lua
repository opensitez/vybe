-- vybe-test: lua/utf8_char_advanced/utf8_len_range
-- origin: languages/lua/tests/lua/test_utf8_char_advanced.rs

local __w1 = "2"
local __i = 0

local s = "Aα😀"
do local __t = tostring(utf8.len(s, 1, 3)); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
