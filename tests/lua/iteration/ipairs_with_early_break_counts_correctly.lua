-- vybe-test: lua/iteration/ipairs_with_early_break_counts_correctly
-- origin: languages/lua/tests/lua/test_iteration.rs

local __w1 = "3"
local __i = 0

local count = 0
for i, v in ipairs({10, 20, 30, 40, 50}) do
  count = count + 1
  if i == 3 then break end
end
do local __t = tostring(count); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
