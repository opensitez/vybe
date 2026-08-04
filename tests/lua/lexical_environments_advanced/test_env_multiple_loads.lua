-- vybe-test: lua/lexical_environments_advanced/test_env_multiple_loads
-- origin: languages/lua/tests/lua/test_lexical_environments_advanced.rs

local __w1 = "5"
local __i = 0

local env = {}
local f1 = load('x = 5', '', 't', env)
local f2 = load('return x', '', 't', env)
f1()
do local __t = tostring(f2()); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
