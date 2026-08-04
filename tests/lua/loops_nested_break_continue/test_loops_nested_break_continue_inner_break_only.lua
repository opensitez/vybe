-- vybe-test: lua/loops_nested_break_continue/test_loops_nested_break_continue_inner_break_only
-- origin: languages/lua/tests/lua/test_loops_nested_break_continue.rs

local __w1 = "6"
local __i = 0

local total = 0
for outer = 1, 2 do
  for inner = 1, 4 do
    if inner == 3 then break end
    total = total + outer
  end
end
do local __t = tostring(total); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
