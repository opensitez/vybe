-- vybe-test: lua/coroutine_wrap_errors/test_wrap_error_number_payload
-- origin: languages/lua/tests/lua/test_coroutine_wrap_errors.rs

local __w1 = "true"
local __i = 0

local f = coroutine.wrap(function() error(42) end)
local ok, err = pcall(f)
do local __t = tostring(ok == false and err == 42); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
