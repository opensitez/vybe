-- vybe-test: lua/control_if_block_scopes/test_control_if_block_scopes_else_and_blocks
-- origin: languages/lua/tests/lua/test_control_if_block_scopes.rs

local __w1 = "9"
local __i = 0

if false then do local __t = tostring(0); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end else do local a = 4 do local b = 5 do local __t = tostring(a + b); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end end end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
