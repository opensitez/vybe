-- vybe-test: lua/loops_while/while_empty_body
-- origin: languages/lua/tests/lua/test_loops_while.rs

local __w1 = "3"
local __i = 0

local i = 1
while (function() i = i + 1; return i < 3 end)() do end
do local __t = tostring(i); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
