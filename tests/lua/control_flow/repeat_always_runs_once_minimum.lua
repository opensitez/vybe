-- vybe-test: lua/control_flow/repeat_always_runs_once_minimum
-- origin: languages/lua/tests/lua/test_control_flow.rs

local __w1 = "true"
local __i = 0

local ran = false
repeat ran = true until true
do local __t = tostring(ran); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
