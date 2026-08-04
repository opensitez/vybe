-- vybe-test: lua/metamethods_bitwise/bnot_metamethod
-- origin: languages/lua/tests/lua/test_metamethods_bitwise.rs

local __w1 = "255"
local __i = 0

local mt = {__bnot = function(a) return {v = ~a.v & 0xFF} end}
local function W(n) return setmetatable({v=n}, mt) end
do local __t = tostring((~W(0)).v); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
