-- vybe-test: lua/metatables_arithmetic/chained_add_op
-- origin: languages/lua/tests/lua/test_metatables_arithmetic.rs

local __w1 = "6"
local __i = 0

local mt={}
mt.__index=mt
mt.__add=function(a,b) return setmetatable({v=a.v+b.v},mt) end
local W=function(n) return setmetatable({v=n},mt) end
do local __t = tostring(((W(1)+W(2))+W(3)).v); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
