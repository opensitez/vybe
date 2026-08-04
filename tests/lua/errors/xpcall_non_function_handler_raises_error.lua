-- vybe-test: lua/errors/xpcall_non_function_handler_raises_error
-- origin: languages/lua/tests/lua/test_errors.rs

local __w1 = "false"
local __i = 0

local ok, err = pcall(function() xpcall(function() error("fail") end, "not_a_function") end)
do local __t = tostring(ok); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
