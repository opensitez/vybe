-- vybe-test: lua/control_if_local_binding/test_control_if_local_binding_string_concat_in_local_branch
-- origin: languages/lua/tests/lua/test_control_if_local_binding.rs

local __w1 = "ab"
local __i = 0

if true then local p = "a" local q = "b" do local __t = tostring(p .. q); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
