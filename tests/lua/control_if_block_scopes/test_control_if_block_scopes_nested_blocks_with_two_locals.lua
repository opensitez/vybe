-- vybe-test: lua/control_if_block_scopes/test_control_if_block_scopes_nested_blocks_with_two_locals
-- origin: languages/lua/tests/lua/test_control_if_block_scopes.rs

local __w1 = "5"
local __i = 0

if true then do local left = 2 do local right = 3 do local __t = tostring(left + right); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end end end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
