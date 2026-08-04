-- vybe-test: lua/string_format_advanced/format_float_precision
-- origin: languages/lua/tests/lua/test_string_format_advanced.rs

local __w1 = "1.235"
local __i = 0

do local __t = tostring(string.format("%.3f", 1.2345)); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
