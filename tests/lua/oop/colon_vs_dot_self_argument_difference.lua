-- vybe-test: lua/oop/colon_vs_dot_self_argument_difference
-- origin: languages/lua/tests/lua/test_oop.rs

local __w1 = "1"
local __i = 0

local t = {n = 0}
function t.inc(self) self.n = self.n + 1 end
t:inc()
do local __t = tostring(t.n); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
