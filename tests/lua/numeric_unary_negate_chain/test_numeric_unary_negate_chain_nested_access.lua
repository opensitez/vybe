-- vybe-test: lua/numeric_unary_negate_chain/test_numeric_unary_negate_chain_nested_access
-- origin: languages/lua/tests/lua/test_numeric_unary_negate_chain.rs

local __w1 = "-7"
local __i = 0

local t = {u = {v = 7}}; do local __t = tostring(-t.u.v); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
