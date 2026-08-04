-- vybe-test: lua/iteration/repeat_until_reads_updated_local
-- origin: languages/lua/tests/lua/test_iteration.rs

local __w1 = "4"
local __i = 0

local n = 0
repeat n = n + 2 until n >= 4
do local __t = tostring(n); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
