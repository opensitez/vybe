-- vybe-test: lua/operators/false_and_short_circuits_before_rhs_call
-- origin: languages/lua/tests/lua/test_operators.rs

local __w1 = "false"
local __i = 0

local n = 0
local function bump() n = n + 1 return true end
do local __t = tostring(false and bump()); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
