-- vybe-test: lua/control_if_elseif_fallthrough/test_control_if_elseif_fallthrough_not_expression
-- origin: languages/lua/tests/lua/test_control_if_elseif_fallthrough.rs

local __w1 = "ok"
local __i = 0

local x=4; local y=''; if not (x < 0) then y='ok' elseif x == 4 then y='four' else y='no' end; do local __t = tostring(y); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
