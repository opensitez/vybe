-- vybe-test: lua/string_format_float/test_string_format_float_nested
-- origin: languages/lua/tests/lua/test_string_format_float.rs

local __w1 = "true"
local __i = 0

local s = string.format("%f", 1.5714285714285714)
do local __t = tostring(type(s) == "string"); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
