-- vybe-test: lua/goto_nested_scoping/goto_for_loop_break
-- origin: languages/lua/tests/lua/test_goto_nested_scoping.rs

local __w1 = "3"
local __i = 0

local last = 0
for i = 1, 10 do
  if i == 4 then goto done end
  last = i
end
::done::
do local __t = tostring(last); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
