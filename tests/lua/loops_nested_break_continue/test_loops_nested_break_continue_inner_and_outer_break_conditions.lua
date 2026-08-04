-- vybe-test: lua/loops_nested_break_continue/test_loops_nested_break_continue_inner_and_outer_break_conditions
-- origin: languages/lua/tests/lua/test_loops_nested_break_continue.rs

local __w1 = "6"
local __i = 0

local total = 0
for outer = 1, 6 do
  for inner = 1, 6 do
    if outer == 4 then break end
    if inner == 3 then break end
    total = total + 1
  end
  if outer == 4 then break end
end
do local __t = tostring(total); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
