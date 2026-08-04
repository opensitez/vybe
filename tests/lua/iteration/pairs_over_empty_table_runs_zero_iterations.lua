-- vybe-test: lua/iteration/pairs_over_empty_table_runs_zero_iterations
-- origin: languages/lua/tests/lua/test_iteration.rs

local __w1 = "0"
local __i = 0

local count = 0
for _ in pairs({}) do count = count + 1 end
do local __t = tostring(count); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
