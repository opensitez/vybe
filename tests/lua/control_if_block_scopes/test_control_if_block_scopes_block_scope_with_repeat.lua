-- vybe-test: lua/control_if_block_scopes/test_control_if_block_scopes_block_scope_with_repeat
-- origin: languages/lua/tests/lua/test_control_if_block_scopes.rs

local __w1 = "1"
local __i = 0

if true then local x = 0 do repeat x = x + 1 until x >= 1 end do local __t = tostring(x); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
