-- vybe-test: lua/control_flow/if_with_comparison_on_strings
-- origin: languages/lua/tests/lua/test_control_flow.rs

local __w1 = "ordered"
local __i = 0

if "a" < "b" then do local __t = tostring("ordered"); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
