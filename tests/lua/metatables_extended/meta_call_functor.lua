-- vybe-test: lua/metatables_extended/meta_call_functor
-- origin: languages/lua/tests/lua/test_metatables_extended.rs

local __w1 = "15"
local __i = 0

local mt = {__call = function(self, a, b) return a + b end}
local obj = setmetatable({}, mt)
do local __t = tostring(obj(5, 10)); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
