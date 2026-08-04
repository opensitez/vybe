-- vybe-test: lua/loops_for_numeric/numeric_for_exact_endpoint_hit_with_step
-- origin: languages/lua/tests/lua/test_loops_for_numeric.rs

local __w1 = "3"
local __i = 0

local n = 0
for i = 0, 10, 5 do n = n + 1 end
do local __t = tostring(n); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
