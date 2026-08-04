-- vybe-test: lua/oop/class_table_holds_shared_method
-- origin: languages/lua/tests/lua/test_oop.rs

local __w1 = "woof"
local __i = 0

local Dog = {}
function Dog:speak() return "woof" end
local d = setmetatable({}, {__index = Dog})
do local __t = tostring(d:speak()); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
