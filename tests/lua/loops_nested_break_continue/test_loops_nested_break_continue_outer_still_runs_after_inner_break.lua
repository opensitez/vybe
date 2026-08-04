-- vybe-test: lua/loops_nested_break_continue/test_loops_nested_break_continue_outer_still_runs_after_inner_break
-- origin: languages/lua/tests/lua/test_loops_nested_break_continue.rs

local __w1 = "3"
local __i = 0

local total = 0
for outer = 1, 3 do
  local skip = false
  for inner = 1, 3 do
    if inner == 2 then skip = true; break end
    total = total + inner
  end
  if not skip then total = total + 10 end
end
do local __t = tostring(total); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
