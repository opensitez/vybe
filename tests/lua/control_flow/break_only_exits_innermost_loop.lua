-- vybe-test: lua/control_flow/break_only_exits_innermost_loop
-- origin: languages/lua/tests/lua/test_control_flow.rs

local __w1 = "3"
local __i = 0

local sum = 0
for i = 1, 3 do
  for j = 1, 3 do
    if j == 2 then break end
    sum = sum + j
  end
end
do local __t = tostring(sum); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
