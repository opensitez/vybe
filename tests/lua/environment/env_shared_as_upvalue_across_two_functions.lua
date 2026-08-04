-- vybe-test: lua/environment/env_shared_as_upvalue_across_two_functions
-- origin: languages/lua/tests/lua/test_environment.rs

local env = {print = print, n = 0}
local function inc() env.n = env.n + 1 end
local function get() return env.n end
inc(); inc(); inc()
print(get())
