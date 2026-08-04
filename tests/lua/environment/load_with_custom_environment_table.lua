-- vybe-test: lua/environment/load_with_custom_environment_table
-- origin: languages/lua/tests/lua/test_environment.rs

local env = {y = 3, print = print}
local f = load("print(y)", "chunk", "t", env)
f()
