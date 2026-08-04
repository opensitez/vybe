-- vybe-test: lua/control_if_block_scopes/test_control_if_block_scopes_conditional_block_scope
-- origin: languages/lua/tests/lua/test_control_if_block_scopes.rs

local __w1 = "10"
local __i = 0

if true then local base = 3; if base > 1 then do local base = base + 7 do local __t = tostring(base); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end end end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
