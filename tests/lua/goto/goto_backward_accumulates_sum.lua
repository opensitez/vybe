-- vybe-test: lua/goto/goto_backward_accumulates_sum
-- origin: languages/lua/tests/lua/test_goto.rs

local __w1 = "10"
local __i = 0

local total = 0
local i = 1
::again::
total = total + i
i = i + 1
if i <= 4 then goto again end
do local __t = tostring(total); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
