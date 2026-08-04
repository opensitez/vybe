-- vybe-test: lua/loops_repeat_guard/test_loops_repeat_guard_conditional_string
-- origin: languages/lua/tests/lua/test_loops_repeat_guard.rs

local __w1 = "aa"
local __i = 0

local out = ''
repeat out = out .. 'a' until #out > 1
do local __t = tostring(out); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
