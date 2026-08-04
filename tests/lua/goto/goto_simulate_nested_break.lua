-- vybe-test: lua/goto/goto_simulate_nested_break
-- origin: languages/lua/tests/lua/test_goto.rs

local __w1 = "11 12 13 21 "
local __i = 0

local s = ""
for i = 1, 3 do
  for j = 1, 3 do
    if i * j == 4 then goto exit_all end
    s = s .. i .. j .. " "
  end
end
::exit_all::
do local __t = tostring(s); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
