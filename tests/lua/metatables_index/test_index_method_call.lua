-- vybe-test: lua/metatables_index/test_index_method_call
-- origin: languages/lua/tests/lua/test_metatables_index.rs

local __w1 = "5"
local __i = 0

local class={}; function class:get() return self.x end; local obj=setmetatable({x=5}, {__index=class}); do local __t = tostring(obj:get()); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
