-- vybe-test: lua/operators_logical/test_or_short_circuit
-- origin: languages/lua/tests/lua/test_operators_logical.rs

local __w1 = "0"
local __i = 0

local a=0; local _ = true or (function() a=1 return true end)(); do local __t = tostring(a); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
