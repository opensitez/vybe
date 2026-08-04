-- vybe-test: lua/control_if_block_scopes/test_control_if_block_scopes_block_scopes_mixed_types
-- origin: languages/lua/tests/lua/test_control_if_block_scopes.rs

local __w1 = "1"
local __i = 0

if true then local count = 0; do local s = "x"; if s == "x" then count = count + 1 end end do local __t = tostring(count); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
