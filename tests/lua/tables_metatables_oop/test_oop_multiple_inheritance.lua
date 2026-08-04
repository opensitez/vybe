-- vybe-test: lua/tables_metatables_oop/test_oop_multiple_inheritance
-- origin: languages/lua/tests/lua/test_tables_metatables_oop.rs

local __w1 = "1 2"
local __i = 0

local function search(k, plist)
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
do local __t = tostring(obj.a .. ' ' .. obj.b); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
