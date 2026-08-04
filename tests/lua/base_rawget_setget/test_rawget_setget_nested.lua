-- vybe-test: lua/base_rawget_setget/test_rawget_setget_nested
-- origin: languages/lua/tests/lua/test_base_rawget_setget.rs

local __w1 = "true"
local __i = 0

local t = {}
rawset(t, "k", 11)
do local __t = tostring(rawget(t, "k") == 11); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
