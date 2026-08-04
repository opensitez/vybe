-- vybe-test: lua/load_custom_env/load_env_write
-- origin: languages/lua/tests/lua/test_load_custom_env.rs

local __w1 = "55"
local __i = 0

local env = setmetatable({}, {__newindex = function(t, k, v) rawset(t, k, v) end})
load("x = 55", "c", "t", env)()
do local __t = tostring(env.x); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
