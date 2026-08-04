-- vybe-test: lua/environment/local_env_table_shadows_globals_for_chunk
-- origin: languages/lua/tests/lua/test_environment.rs

local _ENV = {x = 9, print = print}
print(x)
