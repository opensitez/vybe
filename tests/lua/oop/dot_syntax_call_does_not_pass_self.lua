-- vybe-test: lua/oop/dot_syntax_call_does_not_pass_self
-- origin: languages/lua/tests/lua/test_oop.rs

local __w1 = "2"
local __i = 0

local t = {v = 1}
function t.bump(self) self.v = self.v + 1 end
t.bump(t)
do local __t = tostring(t.v); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
