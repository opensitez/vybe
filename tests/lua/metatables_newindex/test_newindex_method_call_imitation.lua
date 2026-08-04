-- vybe-test: lua/metatables_newindex/test_newindex_method_call_imitation
-- origin: languages/lua/tests/lua/test_metatables_newindex.rs

local __w1 = "5"
local __i = 0

local class={}; function class:set(k, v) rawset(self, k, v) end; local obj=setmetatable({}, {__newindex=function(t,k,v) class.set(t,k,v) end}); obj.x=5; do local __t = tostring(obj.x); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
