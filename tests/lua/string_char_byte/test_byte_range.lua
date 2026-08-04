-- vybe-test: lua/string_char_byte/test_byte_range
-- origin: languages/lua/tests/lua/test_string_char_byte.rs

local __w1 = "65 66 67"
local __i = 0

local a,b,c = string.byte('ABC', 1, 3); do local __t = tostring(a..' '..b..' '..c); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
