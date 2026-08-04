-- vybe-test: lua/loops_repeat_until/repeat_break_from_nested_loop_does_not_exit_outer
-- origin: languages/lua/tests/lua/test_loops_repeat_until.rs

local __w1 = "3"
local __i = 0

local i = 0
repeat
  i = i + 1
  for j = 1, 3 do
    if j == 2 then break end
  end
until i == 3
do local __t = tostring(i); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
