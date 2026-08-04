-- vybe-test: lua/control_flow/nested_while_accumulates
-- origin: languages/lua/tests/lua/test_control_flow.rs

local __w1 = "6"
local __i = 0

local sum = 0
local i = 1
while i <= 3 do
  local j = 1
  while j <= 2 do
    sum = sum + 1
    j = j + 1
  end
  i = i + 1
end
do local __t = tostring(sum); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
