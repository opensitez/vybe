-- vybe-test: lua/loops_for_numeric/numeric_for_continue_via_goto
-- origin: languages/lua/tests/lua/test_loops_for_numeric.rs

local __w1 = "9"
local __i = 0

local sum = 0
for i = 1, 6 do
  if i % 2 ~= 0 then sum = sum + i end
end
do local __t = tostring(sum); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
