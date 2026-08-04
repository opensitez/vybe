-- vybe-test: lua/control_if_block_scopes/test_control_if_block_scopes_blocked_assignment_outside
-- origin: languages/lua/tests/lua/test_control_if_block_scopes.rs

local __w1 = "1"
local __i = 0

if true then local x = 1 do local x = 2 end do local __t = tostring(x == 1 and 1 or 0); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
