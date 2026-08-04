-- vybe-test: lua/strings/string_byte_range_returns_varargs
-- origin: languages/lua/tests/lua/test_strings.rs

local __w1 = "65,66"
local __i = 0

local a,b=string.byte("ABC", 1, 2)
do local __t = tostring(a .. "," .. b); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
