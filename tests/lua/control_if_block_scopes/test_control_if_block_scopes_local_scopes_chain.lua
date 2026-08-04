-- vybe-test: lua/control_if_block_scopes/test_control_if_block_scopes_local_scopes_chain
-- origin: languages/lua/tests/lua/test_control_if_block_scopes.rs

local __w1 = "13"
local __i = 0

if true then local a = 1 do local a = 2 do local a = 3 do local __t = tostring(a + 10); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end end do local __t = tostring(a + 20); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end end do local __t = tostring(a + 30); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
