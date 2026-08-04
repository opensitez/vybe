-- vybe-test: lua/coroutine_patterns/coroutine_wrap_error
-- origin: languages/lua/tests/lua/test_coroutine_patterns.rs

local __w1 = "false"
local __i = 0

local f = coroutine.wrap(function() error("crash") end)
local ok = pcall(f)
do local __t = tostring(ok); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
