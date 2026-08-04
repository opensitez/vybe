-- vybe-test: lua/control_flow_super/ctrl_sup_repeat_5
-- origin: languages/lua/tests/lua/test_control_flow_super.rs

local __w1 = "14"
local __i = 0

local n = 0; repeat n = n + 1 until n >= 14; do local __t = tostring(n); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
