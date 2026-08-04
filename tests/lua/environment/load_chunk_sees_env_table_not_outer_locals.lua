-- vybe-test: lua/environment/load_chunk_sees_env_table_not_outer_locals
-- origin: languages/lua/tests/lua/test_environment.rs

local secret = 999
local env = {print = print, secret = 42}
local f = load('print(secret)', 'test', 't', env)
f()
