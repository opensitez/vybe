-- vybe-test: lua/metatables_arithmetic/add_scalar_op
-- origin: languages/lua/tests/lua/test_metatables_arithmetic.rs

local __w1 = "15"
local __i = 0

local mt={__add=function(a,b)
  local va = type(a)=="table" and a.v or a
  local vb = type(b)=="table" and b.v or b
  return {v=va+vb}
end}
local W=function(n) return setmetatable({v=n}, mt) end
do local __t = tostring((W(5)+10).v); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
