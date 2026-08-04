-- vybe-test: lua/metatables_fallback_inheritance/inherit_self_param
-- origin: languages/lua/tests/lua/test_metatables_fallback_inheritance.rs

local __w1 = "hello obj"
local __i = 0

local proto = {
  greet = function(self) return "hello " .. self.name end
}
local obj = setmetatable({name="obj"}, {__index = proto})
do local __t = tostring(obj:greet()); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
