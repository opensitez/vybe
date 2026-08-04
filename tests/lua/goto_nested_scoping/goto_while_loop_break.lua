-- vybe-test: lua/goto_nested_scoping/goto_while_loop_break
-- origin: languages/lua/tests/lua/test_goto_nested_scoping.rs

local __w1 = "3"
local __i = 0

local n = 0
while n < 5 do
  n = n + 1
  if n == 3 then goto exit_loop end
end
::exit_loop::
do local __t = tostring(n); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
