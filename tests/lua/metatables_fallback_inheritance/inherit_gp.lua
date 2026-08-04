-- vybe-test: lua/metatables_fallback_inheritance/inherit_gp
-- origin: languages/lua/tests/lua/test_metatables_fallback_inheritance.rs

local __w1 = "child,70"
local __i = 0

local gp = {name = "gp", age = 70}
local parent = setmetatable({name = "parent"}, {__index = gp})
local child = setmetatable({name = "child"}, {__index = parent})
do local __t = tostring(child.name .. "," .. child.age); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
