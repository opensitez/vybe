-- vybe-test: lua/loops_nested_break_continue/test_loops_nested_break_continue_while_and_for
-- origin: languages/lua/tests/lua/test_loops_nested_break_continue.rs

local __w1 = "26"
local __i = 0

local total = 0
for outer = 1, 3 do
  local n = 0
  while n < 4 do
    n = n + 1
    if n == 3 then break end
    if outer == 2 then total = total + 10 end
    total = total + 1
  end
end
do local __t = tostring(total); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
