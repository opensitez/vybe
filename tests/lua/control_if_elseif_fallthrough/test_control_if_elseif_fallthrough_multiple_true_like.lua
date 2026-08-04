-- vybe-test: lua/control_if_elseif_fallthrough/test_control_if_elseif_fallthrough_multiple_true_like
-- origin: languages/lua/tests/lua/test_control_if_elseif_fallthrough.rs

local __w1 = "big"
local __i = 0

local x=1; local y=''; if x >= 1 then y='big' elseif x > 1 then y='bigger' else y='none' end; do local __t = tostring(y); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
