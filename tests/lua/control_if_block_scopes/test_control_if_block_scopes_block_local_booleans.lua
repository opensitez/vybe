-- vybe-test: lua/control_if_block_scopes/test_control_if_block_scopes_block_local_booleans
-- origin: languages/lua/tests/lua/test_control_if_block_scopes.rs

local __w1 = "true"
local __i = 0

if true then local ok = false do local seen = true do local __t = tostring(ok == false and seen == true and "true" or "false"); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
