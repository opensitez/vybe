-- vybe-test: lua/string_format_float/test_string_format_float_edge_first
-- origin: languages/lua/tests/lua/test_string_format_float.rs

local __w1 = "true"
local __i = 0

local s = string.format("%f", 2.2857142857142856)
do local __t = tostring(type(s) == "string"); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
