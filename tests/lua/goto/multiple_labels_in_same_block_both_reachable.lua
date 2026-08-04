-- vybe-test: lua/goto/multiple_labels_in_same_block_both_reachable
-- origin: languages/lua/tests/lua/test_goto.rs

local __w1 = "11"
local __i = 0

local x = 0
goto first
::first::
x = 1
goto second
::second::
x = x + 10
do local __t = tostring(x); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
