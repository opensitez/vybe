-- vybe-test: lua/control_flow_exhaustive/ctrl_local_scope_shadow
-- origin: languages/lua/tests/lua/test_control_flow_exhaustive.rs

local __w1 = "2\n1"
local __i = 0

local x = 1
do
  local x = 2
  do local __t = tostring(x); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end
end
do local __t = tostring(x); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
