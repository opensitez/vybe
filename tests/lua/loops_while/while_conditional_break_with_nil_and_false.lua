-- vybe-test: lua/loops_while/while_conditional_break_with_nil_and_false
-- origin: languages/lua/tests/lua/test_loops_while.rs

local __w1 = "5"
local __i = 0

local val = false
local sum = 0
while not val do
  sum = sum + 1
  if sum == 3 then val = nil end
  if sum == 5 then break end
end
do local __t = tostring(sum); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
