-- vybe-test: lua/control_if_local_binding/test_control_if_local_binding_local_returned_from_if_body
-- origin: languages/lua/tests/lua/test_control_if_local_binding.rs

local __w1 = "7"
local __i = 0

if true then local n = 0; if true then n = 7 end do local __t = tostring(n); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
