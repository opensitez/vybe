-- vybe-test: lua/numeric_comparison_truth/test_numeric_comparison_truth_chain_with_false
-- origin: languages/lua/tests/lua/test_numeric_comparison_truth.rs

local __w1 = "false"
local __i = 0

do local __t = tostring(1 < 2 and 4 < 3 and 3 < 4); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
