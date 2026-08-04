-- vybe-test: lua/control_if_local_binding/test_control_if_local_binding_elseif_else_selects_else_local
-- origin: languages/lua/tests/lua/test_control_if_local_binding.rs

local __w1 = "8"
local __i = 0

if false then local x = 1 elseif false then local x = 2 else local x = 8 do local __t = tostring(x); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
