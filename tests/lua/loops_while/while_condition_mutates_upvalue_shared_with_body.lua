-- vybe-test: lua/loops_while/while_condition_mutates_upvalue_shared_with_body
-- origin: languages/lua/tests/lua/test_loops_while.rs

local __w1 = "30"
local __i = 0

local x = 10
local function dec() x = x - 1; return x > 5 end
local sum = 0
while dec() do
  sum = sum + x
end
do local __t = tostring(sum); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
