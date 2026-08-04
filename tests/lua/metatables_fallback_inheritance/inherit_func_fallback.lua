-- vybe-test: lua/metatables_fallback_inheritance/inherit_func_fallback
-- origin: languages/lua/tests/lua/test_metatables_fallback_inheritance.rs

local __w1 = "fallback:foo"
local __i = 0

local fallback = function(t, k) return "fallback:" .. k end
local t = setmetatable({}, {__index = fallback})
do local __t = tostring(t.foo); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
