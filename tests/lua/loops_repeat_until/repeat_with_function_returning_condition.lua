-- vybe-test: lua/loops_repeat_until/repeat_with_function_returning_condition
-- origin: languages/lua/tests/lua/test_loops_repeat_until.rs

local __w1 = "5"
local __i = 0

local count = 0
local function should_stop()
  count = count + 1
  return count >= 5
end
repeat
until should_stop()
do local __t = tostring(count); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
