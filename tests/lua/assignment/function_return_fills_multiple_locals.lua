-- vybe-test: lua/assignment/function_return_fills_multiple_locals
-- origin: languages/lua/tests/lua/test_assignment.rs

local __w1 = "30"
local __i = 0

local function pair() return 10, 20 end
local x, y = pair()
do local __t = tostring(x + y); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
