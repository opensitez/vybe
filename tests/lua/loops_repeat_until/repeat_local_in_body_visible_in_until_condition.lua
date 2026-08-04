-- vybe-test: lua/loops_repeat_until/repeat_local_in_body_visible_in_until_condition
-- origin: languages/lua/tests/lua/test_loops_repeat_until.rs

local __w1 = "3"
local __i = 0

local n = 0
repeat
  local limit = 3
  n = n + 1
until n >= limit
do local __t = tostring(n); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
