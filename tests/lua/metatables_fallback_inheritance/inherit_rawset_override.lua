-- vybe-test: lua/metatables_fallback_inheritance/inherit_rawset_override
-- origin: languages/lua/tests/lua/test_metatables_fallback_inheritance.rs

local __w1 = "2\t1"
local __i = 0

local proto = {x = 1}
local obj = setmetatable({}, {__index = proto})
rawset(obj, "x", 2)
do local __t = tostring(obj.x) .. "\t" .. tostring(proto.x); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
