-- vybe-test: lua/goto_advanced/goto_skips
-- origin: languages/lua/tests/lua/test_goto_advanced.rs

local __w1 = "reached"
local __i = 0

goto skip
do local __t = tostring("skipped"); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end
::skip::
do local __t = tostring("reached"); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
