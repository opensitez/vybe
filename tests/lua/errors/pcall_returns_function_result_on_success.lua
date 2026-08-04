-- vybe-test: lua/errors/pcall_returns_function_result_on_success
-- origin: languages/lua/tests/lua/test_errors.rs

local __w1 = "ok"
local __i = 0

local _, v = pcall(function() return "ok" end)
do local __t = tostring(v); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
