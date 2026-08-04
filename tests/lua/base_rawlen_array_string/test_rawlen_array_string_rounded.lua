-- vybe-test: lua/base_rawlen_array_string/test_rawlen_array_string_rounded
-- origin: languages/lua/tests/lua/test_base_rawlen_array_string.rs

local __w1 = "true"
local __i = 0

local t = {a = 1}; do local __t = tostring(rawlen(t) == 0); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
