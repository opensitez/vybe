-- vybe-test: lua/loops_nested_break_continue/test_loops_nested_break_continue_mix_tables
-- origin: languages/lua/tests/lua/test_loops_nested_break_continue.rs

local __w1 = "2"
local __i = 0

local total = 0
for outer = 1, 2 do
  local t = {1,2,3}
  for _, v in ipairs(t) do
    if v == 2 then break end
    total = total + v
  end
end
do local __t = tostring(total); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
