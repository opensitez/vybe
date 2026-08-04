-- vybe-test: lua/environments_ext/test_env_nested_function
-- origin: languages/lua/tests/lua/test_environments_ext.rs

local _ENV = {print=print, a=42}; local function f() return a end; print(f())
