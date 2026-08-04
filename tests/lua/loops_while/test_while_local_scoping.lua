-- vybe-test: lua/loops_while/test_while_local_scoping
-- origin: languages/lua/tests/lua/test_loops_while.rs

local __w1 = "01"
local __i = 0

local s=''; local x=0; while x<2 do local y=x; x=x+1; s=s..y end; do local __t = tostring(s); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
