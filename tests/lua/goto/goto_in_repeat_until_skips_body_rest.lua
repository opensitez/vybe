-- vybe-test: lua/goto/goto_in_repeat_until_skips_body_rest
-- origin: languages/lua/tests/lua/test_goto.rs

local __w1 = "4"
local __i = 0

local n = 0
local count = 0
repeat
  n = n + 1
  if n % 3 == 0 then goto skip end
  count = count + 1
  ::skip::
until n == 6
do local __t = tostring(count); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
