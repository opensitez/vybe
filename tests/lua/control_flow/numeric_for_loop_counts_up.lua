-- vybe-test: lua/control_flow/numeric_for_loop_counts_up
-- origin: languages/lua/tests/lua/test_control_flow.rs

local __w1 = "15"
local __i = 0

local sum = 0
for i = 1, 5 do
  sum = sum + i
end
do local __t = tostring(sum); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
