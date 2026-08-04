-- vybe-test: lua/oop_metatable_patterns/oop_class_methods
-- origin: languages/lua/tests/lua/test_oop_metatable_patterns.rs

local __w1 = "30"
local __i = 0

local Point = {}
Point.__index = Point
function Point.new(x, y)
  return setmetatable({x=x, y=y}, Point)
end
function Point:sum() return self.x + self.y end
local p = Point.new(10, 20)
do local __t = tostring(p:sum()); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
