-- vybe-test: lua/oop_metatable_patterns/oop_dynamic_dispatch
-- origin: languages/lua/tests/lua/test_oop_metatable_patterns.rs

local __w1 = "A\tB"
local __i = 0

local ClassA = {name = "A"}
ClassA.__index = ClassA
local ClassB = {name = "B"}
ClassB.__index = ClassB
local function get_obj(cond)
  if cond then return setmetatable({}, ClassA) else return setmetatable({}, ClassB) end
end
do local __t = tostring(get_obj(true).name) .. "\t" .. tostring(get_obj(false).name); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
