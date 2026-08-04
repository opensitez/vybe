-- vybe-test: lua/loops_while/while_break_with_returned_values_from_block
-- origin: languages/lua/tests/lua/test_loops_while.rs

local __w1 = "5"
local __i = 0

local function f()
  local i = 1
  while true do
    if i == 5 then return i end
    i = i + 1
  end
end
do local __t = tostring(f()); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
