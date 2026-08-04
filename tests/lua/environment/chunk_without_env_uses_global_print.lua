-- vybe-test: lua/environment/chunk_without_env_uses_global_print
-- origin: languages/lua/tests/lua/test_environment.rs

local __w1 = "function"
local __i = 0

do local __t = tostring(type(print)); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
