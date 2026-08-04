-- vybe-test: lua/loops_nested_break_continue/test_loops_nested_break_continue_outer_break_ends_immediately
-- origin: languages/lua/tests/lua/test_loops_nested_break_continue.rs

local __w1 = "6"
local __i = 0

local count = 0
for outer = 1, 5 do
  for inner = 1, 5 do
    if outer == 2 and inner == 2 then break end
    count = count + 1
  end
  if outer == 2 then break end
end
do local __t = tostring(count); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
