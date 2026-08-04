-- vybe-test: lua/oop/factory_constructor_sets_initial_state
-- origin: languages/lua/tests/lua/test_oop.rs

local __w1 = "7"
local __i = 0

local Point = {}
function Point.new(x, y) return {x = x, y = y} end
local p = Point.new(3, 4)
do local __t = tostring(p.x + p.y); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
