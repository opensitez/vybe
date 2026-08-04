-- vybe-test: lua/numeric_unary_negate_chain/test_numeric_unary_negate_chain_single_negation
-- origin: languages/lua/tests/lua/test_numeric_unary_negate_chain.rs

local __w1 = "-8"
local __i = 0

do local __t = tostring(-8); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
