-- vybe-test: lua/globals/global_persists_across_statements
-- origin: languages/lua/tests/lua/test_globals.rs

local __w1 = "3"
local __i = 0

bar = 1
bar = bar + 2
do local __t = tostring(bar); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
