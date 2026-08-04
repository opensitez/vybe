-- vybe-test: lua/loops_nested_break_continue/test_loops_nested_break_continue_deep_loop_state
-- origin: languages/lua/tests/lua/test_loops_nested_break_continue.rs

local __w1 = "125"
local __i = 0

local total = 0
for outer = 1, 3 do
  for inner = 1, 3 do
    for k = 1, 3 do
      if outer == 2 and inner == 2 and k == 2 then total = total + 100; break end
      total = total + 1
    end
  end
end
do local __t = tostring(total); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
