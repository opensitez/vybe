-- vybe-test: lua/lexical_environments_advanced/test_env_shadowing
-- origin: languages/lua/tests/lua/test_lexical_environments_advanced.rs

local _ENV = {print = print}
local a = 1
_ENV.a = 2
print(a)
