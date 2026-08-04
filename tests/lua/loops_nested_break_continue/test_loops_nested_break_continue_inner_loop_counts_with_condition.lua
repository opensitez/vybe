-- vybe-test: lua/loops_nested_break_continue/test_loops_nested_break_continue_inner_loop_counts_with_condition
-- origin: languages/lua/tests/lua/test_loops_nested_break_continue.rs

local __w1 = "4"
local __i = 0

local total = 0
for outer = 1, 3 do
  for inner = 1, 5 do
    if inner == outer then break end
    total = total + inner
  end
end
do local __t = tostring(total); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
