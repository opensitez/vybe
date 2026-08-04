-- vybe-test: lua/environments_ext/test_env_global_var
-- origin: languages/lua/tests/lua/test_environments_ext.rs

local __w1 = "100"
local __i = 0

_ENV.test_env_var = 100; do local __t = tostring(test_env_var); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
