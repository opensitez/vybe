-- vybe-test: lua/math_logical_compare/test_math_logical_compare_edge_first
-- origin: languages/lua/tests/lua/test_math_logical_compare.rs

local __w1 = "true"
local __i = 0

do local __t = tostring(((15 < 16) and (16 > 15) and (15 <= 16) and (16 >= 15))); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
