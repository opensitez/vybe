-- vybe-test: lua/loops_repeat_until/test_repeat_executes_at_least_once
-- origin: languages/lua/tests/lua/test_loops_repeat_until.rs

local __w1 = "10"
local __i = 0

local i=10; local s=''; repeat s=s..i until true; do local __t = tostring(s); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
