-- vybe-test: lua/goto_advanced/goto_continue_for
-- origin: languages/lua/tests/lua/test_goto_advanced.rs

local __w1 = "9"
local __i = 0

local s = 0
for i = 1, 6 do
  if i % 2 == 0 then goto continue end
  s = s + i
  ::continue::
end
do local __t = tostring(s); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
