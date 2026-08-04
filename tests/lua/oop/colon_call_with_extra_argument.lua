-- vybe-test: lua/oop/colon_call_with_extra_argument
-- origin: languages/lua/tests/lua/test_oop.rs

local __w1 = "5"
local __i = 0

local t = {}
function t.add(self, a, b) return a + b end
do local __t = tostring(t:add(2, 3)); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
