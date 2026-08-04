-- vybe-test: lua/metamethods_bitwise/band_metamethod
-- origin: languages/lua/tests/lua/test_metamethods_bitwise.rs

local __w1 = "15"
local __i = 0

local mt = {__band = function(a, b) return setmetatable({v = a.v & b.v}, getmetatable(a)) end}
mt.__index = mt
local function W(n) return setmetatable({v=n}, mt) end
do local __t = tostring(W(0xFF).__band(W(0xFF), W(0x0F)).v); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
