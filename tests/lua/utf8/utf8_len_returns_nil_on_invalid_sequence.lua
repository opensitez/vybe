-- vybe-test: lua/utf8/utf8_len_returns_nil_on_invalid_sequence
-- origin: languages/lua/tests/lua/test_utf8.rs

local __w1 = "nil"
local __i = 0

do local __t = tostring(tostring(utf8.len("\255"))); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
