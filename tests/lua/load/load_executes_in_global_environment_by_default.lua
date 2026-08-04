-- vybe-test: lua/load/load_executes_in_global_environment_by_default
-- origin: languages/lua/tests/lua/test_load.rs

local __w1 = "8"
local __i = 0

load("x = 8")()
do local __t = tostring(x); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
