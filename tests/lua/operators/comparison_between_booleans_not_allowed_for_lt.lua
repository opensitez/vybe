-- vybe-test: lua/operators/comparison_between_booleans_not_allowed_for_lt
-- origin: languages/lua/tests/lua/test_operators.rs

local __w1 = "false"
local __i = 0

local ok = pcall(function() return true < false end)
do local __t = tostring(ok); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
