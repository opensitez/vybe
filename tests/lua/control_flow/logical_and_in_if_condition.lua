-- vybe-test: lua/control_flow/logical_and_in_if_condition
-- origin: languages/lua/tests/lua/test_control_flow.rs

local __w1 = "ok"
local __i = 0

if true and true then do local __t = tostring("ok"); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
