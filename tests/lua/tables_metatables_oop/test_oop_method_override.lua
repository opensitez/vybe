-- vybe-test: lua/tables_metatables_oop/test_oop_method_override
-- origin: languages/lua/tests/lua/test_tables_metatables_oop.rs

local __w1 = "10"
local __i = 0

local Base = {value = 10}
function Base:new(o)
    o = o or {}
    setmetatable(o, self)
    self.__index = self
    return o
end
function Base:get() return self.value end
local Derived = Base:new()
function Derived:get() return self.value * 2 end
local obj = Derived:new({value = 5})
do local __t = tostring(obj:get()); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
