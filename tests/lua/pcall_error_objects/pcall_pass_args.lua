-- vybe-test: lua/pcall_error_objects/pcall_pass_args
-- origin: languages/lua/tests/lua/test_pcall_error_objects.rs

local __w1 = "true\t42"
local __i = 0

local ok, v = pcall(function(x) return x * 2 end, 21)
do local __t = tostring(ok) .. "\t" .. tostring(v); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
