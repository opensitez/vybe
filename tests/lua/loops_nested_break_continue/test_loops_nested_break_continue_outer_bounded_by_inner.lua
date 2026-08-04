-- vybe-test: lua/loops_nested_break_continue/test_loops_nested_break_continue_outer_bounded_by_inner
-- origin: languages/lua/tests/lua/test_loops_nested_break_continue.rs

local __w1 = "0"
local __i = 0

local total = 0
for outer = 1, 5 do
  local hit = false
  for inner = 1, 5 do
    if inner == 4 then hit = true; break end
  end
  if not hit then total = total + 1 end
end
do local __t = tostring(total); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
