-- vybe-test: lua/goto/goto_simulate_continue
-- origin: languages/lua/tests/lua/test_goto.rs

local __w1 = "124"
local __i = 0

local s = ""
for i = 1, 4 do
  if i == 3 then goto skip end
  s = s .. i
  ::skip::
end
do local __t = tostring(s); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
