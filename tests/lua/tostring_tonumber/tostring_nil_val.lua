-- vybe-test: lua/tostring_tonumber/tostring_nil_val
-- origin: languages/lua/tests/lua/test_tostring_tonumber.rs

local __w1 = "nil"
local __i = 0

do local __t = tostring(tostring(nil)); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
