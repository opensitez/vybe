-- vybe-test: lua/oop/explicit_self_with_dot_call_matches_colon
-- origin: languages/lua/tests/lua/test_oop.rs

local __w1 = "9"
local __i = 0

local obj = {v = 1}
function obj:set(x) self.v = x end
obj.set(obj, 9)
do local __t = tostring(obj.v); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
