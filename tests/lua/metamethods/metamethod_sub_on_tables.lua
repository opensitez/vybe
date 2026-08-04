-- vybe-test: lua/metamethods/metamethod_sub_on_tables
-- origin: languages/lua/tests/lua/test_metamethods.rs

local __w1 = "3"
local __i = 0

local mt={__sub=function(a,b) return a.n-b.n end}
local a=setmetatable({n=5},mt)
local b=setmetatable({n=2},mt)
do local __t = tostring((a-b).n); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
