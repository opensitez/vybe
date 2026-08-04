-- vybe-test: lua/environment/pcall_in_env_table_without_error_function
-- origin: languages/lua/tests/lua/test_environment.rs

local env = {print = print, pcall = pcall, error = error}
local f = load('local ok = pcall(function() error("e") end); print(ok)', 'c', 't', env)
f()
