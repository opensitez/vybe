-- vybe-test: lua/control_flow_heavy/ctrl_heavy_while_6
-- origin: languages/lua/tests/lua/test_control_flow_heavy.rs

local __w1 = "6"
local __i = 0

local n = 0; while n < 6 do n = n + 1 end; do local __t = tostring(n); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
