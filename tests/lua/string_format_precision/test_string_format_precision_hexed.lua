-- vybe-test: lua/string_format_precision/test_string_format_precision_hexed
-- origin: languages/lua/tests/lua/test_string_format_precision.rs

local __w1 = "true"
local __i = 0

local s = string.format("%.5f", 1.6667)
do local __t = tostring(string.find(s, "%.") ~= nil or 5 == 0); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
