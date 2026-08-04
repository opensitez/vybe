-- vybe-test: lua/environments_ext/test_env_shadowing
-- origin: languages/lua/tests/lua/test_environments_ext.rs

local _ENV = {print=print, a=42}; print(a)
