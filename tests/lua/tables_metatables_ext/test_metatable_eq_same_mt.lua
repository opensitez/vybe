-- vybe-test: lua/tables_metatables_ext/test_metatable_eq_same_mt
-- origin: languages/lua/tests/lua/test_tables_metatables_ext.rs

local __w1 = "true"
local __i = 0

local mt={__eq=function(a,b) return a.x==b.x end}; local t1={x=1}; setmetatable(t1, mt); local t2={x=1}; setmetatable(t2, mt); do local __t = tostring(tostring(t1==t2)); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
