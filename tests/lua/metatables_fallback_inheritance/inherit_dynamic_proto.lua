-- vybe-test: lua/metatables_fallback_inheritance/inherit_dynamic_proto
-- origin: languages/lua/tests/lua/test_metatables_fallback_inheritance.rs

local __w1 = "1\n2"
local __i = 0

local protoA = {x = 1}
local protoB = {x = 2}
local mt = {__index = protoA}
local obj = setmetatable({}, mt)
do local __t = tostring(obj.x); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end
mt.__index = protoB
do local __t = tostring(obj.x); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
