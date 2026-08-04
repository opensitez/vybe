-- vybe-test: lua/loops_repeat_guard/test_loops_repeat_guard_guard_with_nil_result
-- origin: languages/lua/tests/lua/test_loops_repeat_guard.rs

local __w1 = "1"
local __i = 0

local done = 0
repeat done = done + 1; local v = nil until v == nil
do local __t = tostring(done); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
