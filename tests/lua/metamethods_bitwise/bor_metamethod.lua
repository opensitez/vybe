-- vybe-test: lua/metamethods_bitwise/bor_metamethod
-- origin: languages/lua/tests/lua/test_metamethods_bitwise.rs

local __w1 = "3"
local __i = 0

local mt = {__bor = function(a, b) return {v = a.v | b.v} end}
local function W(n) return setmetatable({v=n}, mt) end
local r = W(0x01) | W(0x02)
do local __t = tostring(r.v); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
