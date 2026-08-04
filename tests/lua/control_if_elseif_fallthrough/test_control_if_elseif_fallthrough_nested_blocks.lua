-- vybe-test: lua/control_if_elseif_fallthrough/test_control_if_elseif_fallthrough_nested_blocks
-- origin: languages/lua/tests/lua/test_control_if_elseif_fallthrough.rs

local __w1 = "not_one"
local __i = 0

local x=2; local y=''; if x > 1 then if x == 1 then y='one' else y='not_one' end else y='low' end; do local __t = tostring(y); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
