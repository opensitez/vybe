-- vybe-test: lua/loops_repeat_until/repeat_goto_skips_rest_of_body
-- origin: languages/lua/tests/lua/test_loops_repeat_until.rs

local __w1 = "9"
local __i = 0

local sum = 0
local i = 0
repeat
  i = i + 1
  if i % 2 ~= 0 then sum = sum + i end
until i == 6
do local __t = tostring(sum); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
