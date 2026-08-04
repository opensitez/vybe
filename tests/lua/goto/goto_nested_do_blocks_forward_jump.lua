-- vybe-test: lua/goto/goto_nested_do_blocks_forward_jump
-- origin: languages/lua/tests/lua/test_goto.rs

local __w1 = "0"
local __i = 0

local x = 0
do
  do
    goto dest
  end
end
x = 100
::dest::
do local __t = tostring(x); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
