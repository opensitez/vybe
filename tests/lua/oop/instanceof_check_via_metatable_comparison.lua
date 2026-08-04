-- vybe-test: lua/oop/instanceof_check_via_metatable_comparison
-- origin: languages/lua/tests/lua/test_oop.rs

local __w1 = "true"
local __i = 0

local MyClass = {}
MyClass.__index = MyClass
function MyClass.new() return setmetatable({}, MyClass) end
local obj = MyClass.new()
do local __t = tostring(getmetatable(obj) == MyClass); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
