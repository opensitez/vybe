-- vybe-test: lua/pcall_error_objects/nested_pcall_inner
-- origin: languages/lua/tests/lua/test_pcall_error_objects.rs

local __w1 = "false"
local __i = 0

local outer_ok = pcall(function()
  local inner_ok = pcall(function() error("inner") end)
  do local __t = tostring(inner_ok); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end
end)
do local __t = tostring(outer_ok); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
