-- vybe-test: lua/goto_nested_scoping/goto_nested_do_exit
-- origin: languages/lua/tests/lua/test_goto_nested_scoping.rs

local __w1 = "false"
local __i = 0

local reached = false
do
  do
    goto target
  end
  reached = true
end
::target::
do local __t = tostring(reached); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
