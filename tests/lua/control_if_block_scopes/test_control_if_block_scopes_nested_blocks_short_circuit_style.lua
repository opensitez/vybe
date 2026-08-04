-- vybe-test: lua/control_if_block_scopes/test_control_if_block_scopes_nested_blocks_short_circuit_style
-- origin: languages/lua/tests/lua/test_control_if_block_scopes.rs

local __w1 = "1"
local __i = 0

if true then local value = 1 do if true then do local value = 8; value = value + 1 end end do local __t = tostring(value); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
