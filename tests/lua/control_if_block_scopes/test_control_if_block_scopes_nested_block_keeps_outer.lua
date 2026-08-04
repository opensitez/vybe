-- vybe-test: lua/control_if_block_scopes/test_control_if_block_scopes_nested_block_keeps_outer
-- origin: languages/lua/tests/lua/test_control_if_block_scopes.rs

local __w1 = "7"
local __i = 0

if true then local sum = 0; do local x = 2; sum = sum + x; do local x = 5 sum = sum + x end end do local __t = tostring(sum); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
