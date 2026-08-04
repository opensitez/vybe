-- vybe-test: lua/errors/pcall_invokes_function_with_arguments
-- origin: languages/lua/tests/lua/test_errors.rs

local __w1 = "12"
local __i = 0

local _, v = pcall(function(x) return x * 2 end, 6)
do local __t = tostring(v); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
