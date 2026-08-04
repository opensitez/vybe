-- vybe-test: lua/metatables/metatable_call_metamethod_on_table
-- origin: languages/lua/tests/lua/test_metatables.rs

local __w1 = "7"
local __i = 0

local callable = setmetatable({}, {
  __call = function(self, a, b) return a + b end
})
do local __t = tostring(callable(3, 4)); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
