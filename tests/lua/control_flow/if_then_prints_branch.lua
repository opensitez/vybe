-- vybe-test: lua/control_flow/if_then_prints_branch
-- origin: languages/lua/tests/lua/test_control_flow.rs

local __w1 = "yes"
local __i = 0

if 1 < 2 then do local __t = tostring("yes"); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
