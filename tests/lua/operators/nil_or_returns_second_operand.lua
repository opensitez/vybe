-- vybe-test: lua/operators/nil_or_returns_second_operand
-- origin: languages/lua/tests/lua/test_operators.rs

local __w1 = "fallback"
local __i = 0

do local __t = tostring(nil or "fallback"); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
