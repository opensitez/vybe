-- vybe-test: lua/control_if_elseif_fallthrough/test_control_if_elseif_fallthrough_default_shadow
-- origin: languages/lua/tests/lua/test_control_if_elseif_fallthrough.rs

local __w1 = "nil"
local __i = 0

local x=''; local y=''; if x == nil then y='nil' elseif x == '' then y='empty' else y='other' end; do local __t = tostring(y); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
