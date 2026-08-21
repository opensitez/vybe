-- vybe-test: lua/pcall_error_objects/pcall_multi_return_success
-- origin: languages/lua/tests/lua/test_pcall_error_objects.rs

local __w1 = "true\t1\t2"
local __i = 0

local ok, a, b = pcall(function() return 1, 2 end)
do local __t = tostring(ok) .. "\t" .. tostring(a) .. "\t" .. tostring(b); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
