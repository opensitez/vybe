-- vybe-test: lua/coercion/arithmetic_coerces_string_to_number
-- origin: languages/lua/tests/lua/test_coercion.rs

local __w1 = "8"
local __i = 0

do local __t = tostring("5" + 3); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
