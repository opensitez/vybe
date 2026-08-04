-- vybe-test: lua/scoping/scoping_custom_env_redirects_global_writes
-- origin: languages/lua/tests/lua/test_scoping.rs

local __w1 = "42"
local __i = 0

local env = {}
local function run_in_env()
  local _ENV = env
  global_var = 42
end
run_in_env()
do local __t = tostring(env.global_var); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
