-- vybe-test: lua/metamethods/metamethod_unm_negates
-- origin: languages/lua/tests/lua/test_metamethods.rs

local __w1 = "-4"
local __i = 0

local mt={__unm=function(a) return {n=-a.n} end}
do local __t = tostring((-setmetatable({n=4},mt)).n); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
