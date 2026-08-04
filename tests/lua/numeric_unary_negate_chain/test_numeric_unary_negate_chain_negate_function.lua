-- vybe-test: lua/numeric_unary_negate_chain/test_numeric_unary_negate_chain_negate_function
-- origin: languages/lua/tests/lua/test_numeric_unary_negate_chain.rs

local __w1 = "-9"
local __i = 0

function v() return 9 end; do local __t = tostring(-v()); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
