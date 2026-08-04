-- vybe-test: lua/pcall_error_objects/pcall_runtime_error_catch
-- origin: languages/lua/tests/lua/test_pcall_error_objects.rs

local __w1 = "false"
local __i = 0

local ok, _ = pcall(function() return ({}):missing() end)
do local __t = tostring(ok); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
