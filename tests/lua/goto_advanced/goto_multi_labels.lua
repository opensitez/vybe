-- vybe-test: lua/goto_advanced/goto_multi_labels
-- origin: languages/lua/tests/lua/test_goto_advanced.rs

local __w1 = "2"
local __i = 0

local step = 1
goto step2
::step1::
step = 10
goto done
::step2::
step = 2
goto done
::done::
do local __t = tostring(step); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
