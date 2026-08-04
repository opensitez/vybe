-- vybe-test: lua/metamethods/metamethod_bnot_on_tables
-- origin: languages/lua/tests/lua/test_metamethods.rs

local __w1 = "-1"
local __i = 0

local mt={__bnot=function(a) return {n=~a.n} end}
do local __t = tostring((~setmetatable({n=0},mt)).n); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
