-- vybe-test: lua/loops_for_numeric/test_for_num_step_eval_once
-- origin: languages/lua/tests/lua/test_loops_for_numeric.rs

local __w1 = "S135"
local __i = 0

local s=''; local function step() s=s..'S'; return 2 end; for i=1,5,step() do s=s..i end; do local __t = tostring(s); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
