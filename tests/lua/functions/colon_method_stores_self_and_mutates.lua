-- vybe-test: lua/functions/colon_method_stores_self_and_mutates
-- origin: languages/lua/tests/lua/test_functions.rs

local __w1 = "2"
local __i = 0

local obj = {n = 0}
function obj:inc() self.n = self.n + 1 end
obj:inc(); obj:inc()
do local __t = tostring(obj.n); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
