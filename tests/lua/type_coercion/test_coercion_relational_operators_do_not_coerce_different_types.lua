-- vybe-test: lua/type_coercion/test_coercion_relational_operators_do_not_coerce_different_types
-- origin: languages/lua/tests/lua/test_type_coercion.rs

local __w1 = "false"
local __i = 0

do local __t = tostring(1 == '1'); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
