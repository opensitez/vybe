-- vybe-test: lua/coroutine_wrap_errors/test_wrap_error_then_recover
-- origin: languages/lua/tests/lua/test_coroutine_wrap_errors.rs

local __w1 = "true"
local __i = 0

local f = coroutine.wrap(function() error("x") end)
local ok, _ = pcall(f)
local ok2, v = pcall(function() return 1 end)
do local __t = tostring(ok2 == true and v == 1); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
