-- vybe-test: lua/loops_for_numeric/test_for_num_expressions_as_bounds
-- origin: languages/lua/tests/lua/test_loops_for_numeric.rs

local __w1 = "246"
local __i = 0

local s=''; local a=1; local b=3; for i=a+1, b*2, b-1 do s=s..i end; do local __t = tostring(s); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
