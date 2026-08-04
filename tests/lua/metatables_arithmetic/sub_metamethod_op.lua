-- vybe-test: lua/metatables_arithmetic/sub_metamethod_op
-- origin: languages/lua/tests/lua/test_metatables_arithmetic.rs

local __w1 = "7"
local __i = 0

local mt={__sub=function(a,b) return {v=a.v-b.v} end}
local W=function(n) return setmetatable({v=n}, mt) end
do local __t = tostring((W(10)-W(3)).v); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
