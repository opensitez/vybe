-- vybe-test: lua/tables_metatables_oop/test_oop_basic_class
-- origin: languages/lua/tests/lua/test_tables_metatables_oop.rs

local __w1 = "bark"
local __i = 0

local Animal = {sound = 'unknown'}
function Animal:new(o)
    o = o or {}
    setmetatable(o, self)
    self.__index = self
    return o
end
function Animal:speak()
    return self.sound
end
local dog = Animal:new({sound = 'bark'})
do local __t = tostring(dog:speak()); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
