-- vybe-test: lua/metamethods/metamethod_unm_negates_custom_number
-- origin: languages/lua/tests/lua/test_metamethods.rs

local __w1 = "-5"
local __i = 0

local mt = {__unm = function(a) return -a.n end}
do local __t = tostring((-setmetatable({n = 5}, mt)).n); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
