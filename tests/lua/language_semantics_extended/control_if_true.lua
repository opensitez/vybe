-- vybe-test: lua/language_semantics_extended/control_if_true
-- origin: languages/lua/tests/lua/test_language_semantics_extended.rs

local __w1 = "true"
local __i = 0

local x = false
if 1 then x = true end
do local __t = tostring(x); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
