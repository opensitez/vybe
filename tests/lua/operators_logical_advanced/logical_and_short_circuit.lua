-- vybe-test: lua/operators_logical_advanced/logical_and_short_circuit
-- origin: languages/lua/tests/lua/test_operators_logical_advanced.rs

local __w1 = "false"
local __i = 0

local called = false
local function rhs() called = true; return true end
local _ = false and rhs()
do local __t = tostring(called); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
