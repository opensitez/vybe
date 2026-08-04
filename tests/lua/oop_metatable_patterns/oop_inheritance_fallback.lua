-- vybe-test: lua/oop_metatable_patterns/oop_inheritance_fallback
-- origin: languages/lua/tests/lua/test_oop_metatable_patterns.rs

local __w1 = "sound"
local __i = 0

local Animal = {}
Animal.__index = Animal
function Animal:speak() return "sound" end
local Dog = setmetatable({}, Animal)
Dog.__index = Dog
local d = setmetatable({}, Dog)
do local __t = tostring(d:speak()); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
