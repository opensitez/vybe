-- vybe-test: lua/operators_exhaustive/op_add_flt
-- origin: languages/lua/tests/lua/test_operators_exhaustive.rs

local __w1 = "4.0"
local __i = 0

do local __t = tostring(1.5 + 2.5); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
