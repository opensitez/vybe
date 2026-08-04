-- vybe-test: lua/loops_while/while_with_complex_condition_short_circuit
-- origin: languages/lua/tests/lua/test_loops_while.rs

local __w1 = "1,1"
local __i = 0

local i, count = 1, 0
while i <= 5 and (function() count = count + 1; return i % 2 == 0 end)() do
  i = i + 1
end
do local __t = tostring(i .. "," .. count); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
