-- vybe-test: lua/coroutines_exhaustive/co_exh_yield_main_err
-- origin: languages/lua/tests/lua/test_coroutines_exhaustive.rs

local __w1 = "false"
local __i = 0

local ok = pcall(coroutine.yield)
do local __t = tostring(ok); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
