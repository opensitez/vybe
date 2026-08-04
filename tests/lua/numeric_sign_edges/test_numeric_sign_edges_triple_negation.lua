-- vybe-test: lua/numeric_sign_edges/test_numeric_sign_edges_triple_negation
-- origin: languages/lua/tests/lua/test_numeric_sign_edges.rs

local __w1 = "-12"
local __i = 0

do local __t = tostring(-(-(-12))); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
