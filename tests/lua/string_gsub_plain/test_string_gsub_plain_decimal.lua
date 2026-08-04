-- vybe-test: lua/string_gsub_plain/test_string_gsub_plain_decimal
-- origin: languages/lua/tests/lua/test_string_gsub_plain.rs

local __w1 = "true"
local __i = 0

local s = string.rep("x+y", 4)
local _, replaced = string.gsub(s, "x+y", "z", 1)
do local __t = tostring(replaced == 1); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
