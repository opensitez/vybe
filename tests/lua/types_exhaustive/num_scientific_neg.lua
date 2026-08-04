-- vybe-test: lua/types_exhaustive/num_scientific_neg
-- origin: languages/lua/tests/lua/test_types_exhaustive.rs

local __w1 = "0.01"
local __i = 0

do local __t = tostring(1e-2); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
