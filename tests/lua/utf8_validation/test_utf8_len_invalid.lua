-- vybe-test: lua/utf8_validation/test_utf8_len_invalid
-- origin: languages/lua/tests/lua/test_utf8_validation.rs

local __w1 = "nil 2"
local __i = 0

local len, pos = utf8.len('a\xFFb'); do local __t = tostring(tostring(len)..' '..pos); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
