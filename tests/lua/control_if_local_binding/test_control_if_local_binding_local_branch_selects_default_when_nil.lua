-- vybe-test: lua/control_if_local_binding/test_control_if_local_binding_local_branch_selects_default_when_nil
-- origin: languages/lua/tests/lua/test_control_if_local_binding.rs

local __w1 = "nil"
local __i = 0

if false then local value = 1 else local value = nil do local __t = tostring(value == nil and "nil" or "v"); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
