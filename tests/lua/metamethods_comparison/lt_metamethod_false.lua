-- vybe-test: lua/metamethods_comparison/lt_metamethod_false
-- origin: languages/lua/tests/lua/test_metamethods_comparison.rs

local __w1 = "false"
local __i = 0

local mt = {__lt = function(a, b) return a.v < b.v end}
local function W(n) return setmetatable({v=n}, mt) end
do local __t = tostring(W(5) < W(2)); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
