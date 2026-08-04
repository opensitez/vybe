-- vybe-test: lua/goto/goto_jumps_from_inside_if_to_after_end
-- origin: languages/lua/tests/lua/test_goto.rs

local __w1 = "1"
local __i = 0

local x = 0
if true then
  x = 1
  goto after
  x = 2
end
::after::
do local __t = tostring(x); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
