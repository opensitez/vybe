-- vybe-test: lua/loops_while/while_nested_control_with_labels_and_gotos
-- origin: languages/lua/tests/lua/test_loops_while.rs

local __w1 = "132"
local __i = 0

local i = 0
local sum = 0
while i < 3 do
  i = i + 1
  local j = 0
  while j < 3 do
    j = j + 1
    if j ~= 2 then sum = sum + (i * 10 + j) end
  end
end
do local __t = tostring(sum); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
