-- vybe-test: lua/string_format/format_float_negative_zero_shows_sign_with_plus
-- origin: languages/lua/tests/lua/test_string_format.rs

local __w1 = "true"
local __i = 0

do local __t = tostring(string.format("%+f", -0.0) == "-0.000000" or string.format("%+f", -0.0) == "+0.000000"); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
