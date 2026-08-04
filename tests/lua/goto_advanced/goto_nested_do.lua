-- vybe-test: lua/goto_advanced/goto_nested_do
-- origin: languages/lua/tests/lua/test_goto_advanced.rs

local __w1 = "out"
local __i = 0

do
  do
    goto out
  end
  do local __t = tostring("inner"); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end
end
::out::
do local __t = tostring("out"); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
