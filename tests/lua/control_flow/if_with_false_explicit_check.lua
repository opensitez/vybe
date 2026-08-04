-- vybe-test: lua/control_flow/if_with_false_explicit_check
-- origin: languages/lua/tests/lua/test_control_flow.rs

local __w1 = "off"
local __i = 0

local flag = false
if flag == false then do local __t = tostring("off"); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
