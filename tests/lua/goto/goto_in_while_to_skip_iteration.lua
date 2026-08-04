-- vybe-test: lua/goto/goto_in_while_to_skip_iteration
-- origin: languages/lua/tests/lua/test_goto.rs

local __w1 = "9"
local __i = 0

local sum = 0
local i = 0
while i < 6 do
  i = i + 1
  if i % 2 == 0 then goto next end
  sum = sum + i
  ::next::
end
do local __t = tostring(sum); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
