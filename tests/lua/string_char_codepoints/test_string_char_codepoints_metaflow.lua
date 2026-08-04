-- vybe-test: lua/string_char_codepoints/test_string_char_codepoints_metaflow
-- origin: languages/lua/tests/lua/test_string_char_codepoints.rs

local __w1 = "true"
local __i = 0

local c = string.char(76); do local __t = tostring(string.byte(c) == 76); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
