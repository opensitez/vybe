-- vybe-test: lua/oop_metatable_patterns/oop_class_instantiation
-- origin: languages/lua/tests/lua/test_oop_metatable_patterns.rs

local __w1 = "3\t4"
local __i = 0

local Point = {}
Point.__index = Point
function Point.new(x, y)
  return setmetatable({x=x, y=y}, Point)
end
local p = Point.new(3, 4)
do local __t = tostring(p.x) .. "\t" .. tostring(p.y); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
