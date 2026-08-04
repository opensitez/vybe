-- vybe-test: lua/coroutines_extended/co_running_main
-- origin: languages/lua/tests/lua/test_coroutines_extended.rs

local __w1 = "thread\ttrue"
local __i = 0

local running, is_main = coroutine.running()
do local __t = tostring(type(running)) .. "\t" .. tostring(is_main); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
