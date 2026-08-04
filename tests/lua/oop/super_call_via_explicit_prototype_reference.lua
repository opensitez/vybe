-- vybe-test: lua/oop/super_call_via_explicit_prototype_reference
-- origin: languages/lua/tests/lua/test_oop.rs

local __w1 = "generic+woof"
local __i = 0

local Animal = {}
function Animal:sound() return 'generic' end
local Dog = setmetatable({}, {__index = Animal})
function Dog:sound() return Animal.sound(self) .. '+woof' end
local d = setmetatable({}, {__index = Dog})
do local __t = tostring(d:sound()); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
