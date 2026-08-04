-- vybe-test: lua/utf8_ext/test_utf8_offset_negative_n
-- origin: languages/lua/tests/lua/test_utf8_ext.rs

local __w1 = "1"
local __i = 0

do local __t = tostring(utf8.offset('你好', -1)); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
