-- vybe-test: lua/control_if_block_scopes/test_control_if_block_scopes_function_calls_inside_block
-- origin: languages/lua/tests/lua/test_control_if_block_scopes.rs

local __w1 = "4"
local __i = 0

if true then local f = function(v) return v + 1 end do do local __t = tostring(f(3)); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
