-- vybe-test: lua/string_format/format_scientific_lowercase
-- origin: languages/lua/tests/lua/test_string_format.rs

local __w1 = "1.000000e+03"
local __i = 0

do local __t = tostring(string.format("%e", 1000)); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
