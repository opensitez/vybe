-- vybe-test: lua/control_flow/if_true_picks_then_over_else
-- origin: languages/lua/tests/lua/test_control_flow.rs

local __w1 = "t"
local __i = 0

if true then do local __t = tostring("t"); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end else do local __t = tostring("f"); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
