-- vybe-test: lua/operators_logical_advanced/logical_ternary_false
-- origin: languages/lua/tests/lua/test_operators_logical_advanced.rs

local __w1 = "no"
local __i = 0

local condition = false
local val = condition and "yes" or "no"
do local __t = tostring(val); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
