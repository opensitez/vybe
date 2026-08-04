-- vybe-test: lua/control_flow/else_runs_when_all_branches_false
-- origin: languages/lua/tests/lua/test_control_flow.rs

local __w1 = "3"
local __i = 0

if false then do local __t = tostring(1); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end elseif false then do local __t = tostring(2); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end else do local __t = tostring(3); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
