lua_print! {
    test_oop_basic_class => {
        "local Animal = {sound = 'unknown'}
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
print(dog:speak())",
        "bark"
    },
    test_oop_inheritance => {
        "local Animal = {sound = 'unknown'}
function Animal:new(o)
    o = o or {}
    setmetatable(o, self)
    self.__index = self
    return o
end
function Animal:speak()
    return self.sound
end
local Dog = Animal:new({sound = 'bark'})
local my_dog = Dog:new()
print(my_dog:speak())",
        "bark"
    },
    test_oop_method_override => {
        "local Base = {value = 10}
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
print(obj:get())",
        "10"
    },
    test_oop_super_call => {
        "local Base = {}
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
print(obj:init())",
        "base sub"
    },
    test_oop_multiple_inheritance => {
        "local function search(k, plist)
    for i = 1, #plist do
        local v = plist[i][k]
        if v then return v end
    end
end
local function createClass(...)
    local c = {}
    local parents = {...}
    setmetatable(c, {__index = function(t, k)
        return search(k, parents)
    end})
    c.__index = c
    function c:new(o)
        o = o or {}
        setmetatable(o, c)
        return o
    end
    return c
end
local A = {a = 1}
local B = {b = 2}
local C = createClass(A, B)
local obj = C:new()
print(obj.a .. ' ' .. obj.b)",
        "1 2"
    },
    test_oop_privacy_closure => {
        "local function make_account(initial)
    local balance = initial
    return {
        withdraw = function(v)
            balance = balance - v
            return balance
        end,
        deposit = function(v)
            balance = balance + v
            return balance
        end,
        get_balance = function()
            return balance
        end
    }
end
local acc = make_account(100)
acc.withdraw(20)
print(acc.get_balance())",
        "80"
    }
}
