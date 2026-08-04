-- vybe-test: lua/control_if_local_binding/test_control_if_local_binding_outer_block_stays_outside_shadow
-- origin: languages/lua/tests/lua/test_control_if_local_binding.rs

local __w1 = "5"
local __i = 0

if true then local a = 5; do local a = 7 end do local __t = tostring(a); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
