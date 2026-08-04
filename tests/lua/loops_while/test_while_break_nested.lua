-- vybe-test: lua/loops_while/test_while_break_nested
-- origin: languages/lua/tests/lua/test_loops_while.rs

local __w1 = "112131"
local __i = 0

local i=1; local s=''; while i<=3 do local j=1; while j<=3 do if j==2 then break end; s=s..i..j; j=j+1 end; i=i+1 end; do local __t = tostring(s); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
