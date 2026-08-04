-- vybe-test: lua/metatables_math/test_math_add_tables
-- origin: languages/lua/tests/lua/test_metatables_math.rs

local __w1 = "30"
local __i = 0

local t1={v=10}; local t2={v=20}; local mt={__add=function(a,b) return a.v+b.v end}; setmetatable(t1, mt); setmetatable(t2, mt); do local __t = tostring(t1+t2); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
