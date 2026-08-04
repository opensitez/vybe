-- vybe-test: lua/string_format_specifiers/format_char
-- origin: languages/lua/tests/lua/test_string_format_specifiers.rs

local __w1 = "A"
local __i = 0

do local __t = tostring(string.format("%c", 65)); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
