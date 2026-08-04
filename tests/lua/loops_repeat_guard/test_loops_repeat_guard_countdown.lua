-- vybe-test: lua/loops_repeat_guard/test_loops_repeat_guard_countdown
-- origin: languages/lua/tests/lua/test_loops_repeat_guard.rs

local __w1 = "3"
local __i = 0

local n = 3
local out = 0
repeat out = out + 1; n = n - 1 until n == 0
do local __t = tostring(out); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
