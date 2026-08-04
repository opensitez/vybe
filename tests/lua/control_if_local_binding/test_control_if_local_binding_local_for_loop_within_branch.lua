-- vybe-test: lua/control_if_local_binding/test_control_if_local_binding_local_for_loop_within_branch
-- origin: languages/lua/tests/lua/test_control_if_local_binding.rs

local __w1 = "6"
local __i = 0

if true then local total = 0; for i = 1, 3 do total = total + i end do local __t = tostring(total); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
