-- vybe-test: lua/goto_advanced/goto_if_end
-- origin: languages/lua/tests/lua/test_goto_advanced.rs

local __w1 = "5"
local __i = 0

local x = 5
if x > 3 then goto done end
do local __t = tostring("not reached"); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end
::done::
do local __t = tostring(x); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
