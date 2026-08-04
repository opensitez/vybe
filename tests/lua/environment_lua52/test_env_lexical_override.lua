-- vybe-test: lua/environment_lua52/test_env_lexical_override
-- origin: languages/lua/tests/lua/test_environment_lua52.rs

local __w1 = "1"
local __i = 0

local a=1; local _ENV={a=2}; do local __t = tostring(a); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
