-- vybe-test: lua/string_byte_indices/test_string_byte_indices_offset
-- origin: languages/lua/tests/lua/test_string_byte_indices.rs

local __w1 = "true"
local __i = 0

do local __t = tostring(string.byte("abcdefghijklmnopqrstuvwxyz", 9) == 105); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
