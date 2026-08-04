-- vybe-test: lua/control_flow/if_false_leaves_local_unchanged
-- origin: languages/lua/tests/lua/test_control_flow.rs

local __w1 = "1"
local __i = 0

local x = 1
if false then x = 2 end
do local __t = tostring(x); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
