-- vybe-test: lua/loops_repeat_until/repeat_nested_outer_break_exits_correctly
-- origin: languages/lua/tests/lua/test_loops_repeat_until.rs

local __w1 = "3"
local __i = 0

local result = 0
local i = 0
repeat
  i = i + 1
  repeat
    result = result + 1
    break
  until true
until i == 3
do local __t = tostring(result); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
