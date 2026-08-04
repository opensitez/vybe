-- vybe-test: lua/environment/load_sets_env_upvalue_in_chunk
-- origin: languages/lua/tests/lua/test_environment.rs

local env = {x = 10, print = print}
local chunk = load('x = x + 5; print(x)', 'c', 't', env)
chunk()
print(env.x)
