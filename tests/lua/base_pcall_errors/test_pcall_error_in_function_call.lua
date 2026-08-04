-- vybe-test: lua/base_pcall_errors/test_pcall_error_in_function_call
-- origin: languages/lua/tests/lua/test_base_pcall_errors.rs

local __w1 = "true"
local __i = 0

local ok, fn = pcall(function() return function() return 1 end end)
do local __t = tostring(ok and type(fn) == "function"); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
