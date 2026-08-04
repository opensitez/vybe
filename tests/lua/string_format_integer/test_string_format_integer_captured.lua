-- vybe-test: lua/string_format_integer/test_string_format_integer_captured
-- origin: languages/lua/tests/lua/test_string_format_integer.rs

local __w1 = "true"
local __i = 0

do local __t = tostring(string.format("%d", 15) == "15"); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
