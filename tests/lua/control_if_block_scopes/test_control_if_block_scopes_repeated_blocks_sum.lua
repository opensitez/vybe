-- vybe-test: lua/control_if_block_scopes/test_control_if_block_scopes_repeated_blocks_sum
-- origin: languages/lua/tests/lua/test_control_if_block_scopes.rs

local __w1 = "3"
local __i = 0

if true then local sum = 0; do local v = 1 sum = sum + v end do local v = 2 sum = sum + v end do local __t = tostring(sum); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
