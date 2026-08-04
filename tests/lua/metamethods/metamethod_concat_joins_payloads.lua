-- vybe-test: lua/metamethods/metamethod_concat_joins_payloads
-- origin: languages/lua/tests/lua/test_metamethods.rs

local __w1 = "ab"
local __i = 0

local mt={__concat=function(a,b) return a.s..b.s end}
do local __t = tostring(setmetatable({s="a"},mt)..setmetatable({s="b"},mt)); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
