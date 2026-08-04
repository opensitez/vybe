-- vybe-test: lua/control_flow/if_uses_local_in_condition
-- origin: languages/lua/tests/lua/test_control_flow.rs

local __w1 = "big"
local __i = 0

local n = 3
if n > 2 then do local __t = tostring("big"); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
