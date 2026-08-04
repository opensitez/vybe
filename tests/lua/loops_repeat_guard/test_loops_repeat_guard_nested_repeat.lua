-- vybe-test: lua/loops_repeat_guard/test_loops_repeat_guard_nested_repeat
-- origin: languages/lua/tests/lua/test_loops_repeat_guard.rs

local __w1 = "9"
local __i = 0

local n = 0
local total = 0
repeat
  n = n + 1
  local inner = 0
  repeat
    inner = inner + 1
    total = total + inner
  until inner > 1
until n > 2
do local __t = tostring(total); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
