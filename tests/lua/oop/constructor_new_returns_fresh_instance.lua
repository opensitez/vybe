-- vybe-test: lua/oop/constructor_new_returns_fresh_instance
-- origin: languages/lua/tests/lua/test_oop.rs

local __w1 = "0"
local __i = 0

local Counter = {}
function Counter.new()
  return {n = 0}
end
local c = Counter.new()
do local __t = tostring(c.n); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
