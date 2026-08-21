-- vybe-test: lua/type_coercion/test_coercion_string_to_number_add
-- origin: languages/lua/tests/lua/test_type_coercion.rs

local __w1 = "15"
local __i = 0

do local __t = tostring('10' + 5); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
