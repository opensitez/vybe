-- vybe-test: lua/string_char_byte/test_byte_out_of_bounds
-- origin: languages/lua/tests/lua/test_string_char_byte.rs

local __w1 = "nil"
local __i = 0

local b = string.byte('ABC', 4); do local __t = tostring(tostring(b)); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
