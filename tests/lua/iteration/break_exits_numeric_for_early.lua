-- vybe-test: lua/iteration/break_exits_numeric_for_early
-- origin: languages/lua/tests/lua/test_iteration.rs

local __w1 = "10"
local __i = 0

local sum=0
for i=1,100 do
  if i>4 then break end
  sum=sum+i
end
do local __t = tostring(sum); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
