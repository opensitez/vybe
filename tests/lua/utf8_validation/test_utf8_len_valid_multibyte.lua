-- vybe-test: lua/utf8_validation/test_utf8_len_valid_multibyte
-- origin: languages/lua/tests/lua/test_utf8_validation.rs

local __w1 = "2"
local __i = 0

do local __t = tostring(utf8.len('你好')); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
