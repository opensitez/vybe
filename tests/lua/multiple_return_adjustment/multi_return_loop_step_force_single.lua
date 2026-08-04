-- vybe-test: lua/multiple_return_adjustment/multi_return_loop_step_force_single
-- origin: languages/lua/tests/lua/test_multiple_return_adjustment.rs

local __w1 = "12"
local __i = 0

local function step() return 2, 99 end
local s = 0
for i = 1, 6, (step()) do s = s + i end
do local __t = tostring(s); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
