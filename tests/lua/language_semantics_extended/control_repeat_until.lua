-- vybe-test: lua/language_semantics_extended/control_repeat_until
-- origin: languages/lua/tests/lua/test_language_semantics_extended.rs

local __w1 = "3"
local __i = 0

local n = 0
repeat n = n + 1 until n >= 3
do local __t = tostring(n); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
