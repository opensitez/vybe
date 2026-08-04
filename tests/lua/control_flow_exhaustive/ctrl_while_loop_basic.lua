-- vybe-test: lua/control_flow_exhaustive/ctrl_while_loop_basic
-- origin: languages/lua/tests/lua/test_control_flow_exhaustive.rs

local __w1 = "3"
local __i = 0

local n = 0; while n < 3 do n = n + 1 end; do local __t = tostring(n); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
