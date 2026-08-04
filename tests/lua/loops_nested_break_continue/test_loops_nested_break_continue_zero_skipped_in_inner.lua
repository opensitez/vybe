-- vybe-test: lua/loops_nested_break_continue/test_loops_nested_break_continue_zero_skipped_in_inner
-- origin: languages/lua/tests/lua/test_loops_nested_break_continue.rs

local __w1 = "13"
local __i = 0

local total = 0
for outer = 1, 4 do
  local n = 0
  repeat
    n = n + 1
    if n == 1 then total = total + 1 else total = total + 2 end
  until n == 2
  if outer == 2 then total = total + 1 end
end
do local __t = tostring(total); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
