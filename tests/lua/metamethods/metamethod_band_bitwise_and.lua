-- vybe-test: lua/metamethods/metamethod_band_bitwise_and
-- origin: languages/lua/tests/lua/test_metamethods.rs

local __w1 = "3"
local __i = 0

local mt = {__band = function(a, b) return {n = a.n & b.n} end}
do local __t = tostring((setmetatable({n=0xF}, mt) & setmetatable({n=0x3}, mt)).n); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
