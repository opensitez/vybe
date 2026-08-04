-- vybe-test: lua/goto/goto_backward_reexecutes_local_initialization
-- origin: languages/lua/tests/lua/test_goto.rs

local __w1 = "13"
local __i = 0

local count = 0
::start::
local x = 10
count = count + 1
x = x + count
if count < 3 then goto start end
do local __t = tostring(x); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
