-- vybe-test: lua/metatables_math/test_math_string_metamethod
-- origin: languages/lua/tests/lua/test_metatables_math.rs

local __w1 = "ab"
local __i = 0

debug.setmetatable('', {__add=function(a,b) return a..b end}); do local __t = tostring('a'+'b'); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
