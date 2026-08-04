-- vybe-test: lua/control_if_local_binding/test_control_if_local_binding_elseif_keeps_that_branch
-- origin: languages/lua/tests/lua/test_control_if_local_binding.rs

local __w1 = "3"
local __i = 0

if false then local value = 1 elseif true then local value = 3 do local __t = tostring(value); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
