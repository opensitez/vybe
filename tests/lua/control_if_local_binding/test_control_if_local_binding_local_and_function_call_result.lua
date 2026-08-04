-- vybe-test: lua/control_if_local_binding/test_control_if_local_binding_local_and_function_call_result
-- origin: languages/lua/tests/lua/test_control_if_local_binding.rs

local __w1 = "8"
local __i = 0

if true then local f = function(v) return v * 2 end do local __t = tostring(f(4)); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
