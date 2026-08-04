-- vybe-test: lua/metatables_arithmetic/add_metamethod_op
-- origin: languages/lua/tests/lua/test_metatables_arithmetic.rs

local __w1 = "7"
local __i = 0

local mt={__add=function(a,b) return setmetatable({v=a.v+b.v},mt) end}
mt.__index=mt
local W=function(n) return setmetatable({v=n},mt) end
do local __t = tostring((W(3)+W(4)).v); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
