-- vybe-test: lua/operators_logical_advanced/logical_coalesce
-- origin: languages/lua/tests/lua/test_operators_logical_advanced.rs

local __w1 = "42"
local __i = 0

local x = nil
local y = 42
do local __t = tostring(x or y or 100); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
