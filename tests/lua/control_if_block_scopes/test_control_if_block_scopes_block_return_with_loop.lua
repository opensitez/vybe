-- vybe-test: lua/control_if_block_scopes/test_control_if_block_scopes_block_return_with_loop
-- origin: languages/lua/tests/lua/test_control_if_block_scopes.rs

local __w1 = "6"
local __i = 0

if true then local total = 0; do for i = 1, 3 do total = total + i end end do local __t = tostring(total); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
