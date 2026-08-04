-- vybe-test: lua/tables_metatables_oop/test_oop_super_call
-- origin: languages/lua/tests/lua/test_tables_metatables_oop.rs

local __w1 = "base sub"
local __i = 0

local Base = {}
function Base:new(o)
    o = o or {}
    setmetatable(o, self)
    self.__index = self
    return o
end
function Base:init() return 'base' end
local Sub = Base:new()
function Sub:init() return Base.init(self) .. ' sub' end
local obj = Sub:new()
do local __t = tostring(obj:init()); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
