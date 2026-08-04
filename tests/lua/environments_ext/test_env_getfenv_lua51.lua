-- vybe-test: lua/environments_ext/test_env_getfenv_lua51
-- origin: languages/lua/tests/lua/test_environments_ext.rs

local __w1 = "false"
local __i = 0

local ok = pcall(function() getfenv(1) end); do local __t = tostring(tostring(ok)); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
