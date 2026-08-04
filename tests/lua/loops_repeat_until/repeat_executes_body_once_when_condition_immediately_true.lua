-- vybe-test: lua/loops_repeat_until/repeat_executes_body_once_when_condition_immediately_true
-- origin: languages/lua/tests/lua/test_loops_repeat_until.rs

local __w1 = "true"
local __i = 0

local executed = false
repeat
  executed = true
until true
do local __t = tostring(executed); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
